/********************************************************
 * 
 * Author:              William Mills
 *                    	Solutions Engineer
 *                    	wimills@cisco.com
 *                    	Cisco Systems
 * 
 * 
 * Version: 1-3-0
 * Released: 03/02/26
 * 
 * Audio Ducking Macro:
 * 
 * This macro monitors the vumeter levels of an incoming mic
 * and automatically ducks (sets a low Gain or Level) of other
 * mics when the monitors mic is considered high.
 * 
 * This is useful when you are using voice lift in a room and 
 * would like to duck any ceiling mics while a person using 
 * a voice lift mic is talking.
 * 
 * v1-2-0 Changes:
 * 
 * Support for MTR Devices added, now joining an MTR call
 * will trigger the macro to begin monitoring the audio.
 * 
 * v1-3-0 Changes:
 * 
 * Added audio event timeouts to handle situations where the
 * monitored audio inputs are muted (externally) which causes
 * no vumeter events and therefore prevents unducking of the mics.
 * 
 * 
 * Full Readme, source code and license details for this macro 
 * are available GitHub:
 * https://github.com/wxsd-sales/audio-ducking-macro
 * 
 ********************************************************/

import xapi from 'xapi';


/*********************************************************
 * Configure the settings below
**********************************************************/

const config = {
  button: {                     // Customise the macros control button name, color and icon
    name: 'Audio Modes',        // Button/Panel/Alert Name
    color: '#f58142',           // Button Color
    icon: 'Sliders',            // Button Icon
    location: 'CallControls'    // Button Location
  },
  showAlerts: true,             // true = show alert messages on controller, false = don't show alerts
  modeNames: {                  // Customise the macros mode names
    autoDuck: 'Auto Adjust Audience',
    presentersOnly: 'Presenters Only',
    presentersAndAudience: 'Presenters & Audience'
  },
  defaultMode: 'autoDuck',      // Specify the default mode
  mics: [                       // Specify which mics should be monitored
    { ConnectorType: 'Microphone', ConnectorId: 1 }   // { ConnectorType: 'Microphone' | 'Ethernet' | 'USBMicrophone'}
  ],
  duck: [                       // Specify which mics should be ducked or unducked
    { ConnectorType: 'Ethernet', ConnectorId: 1, SubId: 1 },
    { ConnectorType: 'Ethernet', ConnectorId: 2, SubId: 1 },
    { ConnectorType: 'Ethernet', ConnectorId: 3, SubId: 1 },
    { ConnectorType: 'Ethernet', ConnectorId: 4, SubId: 1 },
    { ConnectorType: 'Ethernet', ConnectorId: 5, SubId: 1 },
    { ConnectorType: 'Ethernet', ConnectorId: 6, SubId: 1 }
  ],
  threshold: {                  // Specify the thresholds in which the monitors mic is considered high or low
    high: 30,
    low: 25
  },
  levels: {                     // Specify the Gain/Levels which should be set ducked or unducked
    duck: 0,
    unduck: 30
  },
  unduck: {
    timeout: 2                  // Specify the duration where the monitors mic is low before unducking
  },
  samples: 4,                   // The number of samples taken every 100ms, 4 samples at 100ms = 400ms
  debug: false,
  panelId: 'audioDucking'
}


/*********************************************************
 * Do not change below
**********************************************************/


const startVuMeterConnectors = consolidateConnectors(config.mics);
const micNames = createMicStrings(config.mics);
const duck = config.duck.map(({ ConnectorType, ConnectorId, SubId }) => ConnectorType + '.' + ConnectorId + (SubId ? '.' + SubId : ''));
const sampleInterval = 100;
const averageLogFequency = 20;

let gainLevel = 'Gain';
let ducked = false;
let unduckTimer;
let audioEventTimeout;
let listener;
let micLevels;
let micLevelAverages;
let audioEventCount = 0;
let mode;
let callId;


setTimeout(init, 3000);

async function init() {

  gainLevel = await checkGainLevel();

  await createPanel();

  xapi.Event.UserInterface.Extensions.Widget.Action.on(processActions);

  xapi.Status.Call.on(({ ghost, Status, id }) => {
    if (Status && Status == 'Connected' && callId != id) {
      callId = id;
      return processNewCall('RoomOS');
    }

    if (ghost) return processCallEnd('RoomOS');
  });

  xapi.Status.MicrosoftTeams.Calling.InCall.on((inMTRCall) => {
    if (inMTRCall == 'True') return processNewCall('MTR');
    return processCallEnd('MTR');
  });

  const widgets = await xapi.Status.UserInterface.Extensions.Widget.get()
  const selection = widgets.find(widget => widget.WidgetId == config.panelId);
  const value = selection?.Value;

  mode = value && value != '' ? value : config.defaultMode;

  xapi.Command.UserInterface.Extensions.Widget.SetValue({ WidgetId: config.panelId, Value: mode })

  applyMode();

}


async function applyMode() {
  const inCall = await checkInCall();
  if (!inCall) return

  console.log('Applying Mode:', mode);

  if (mode == 'presentersOnly') {
    stopMonitor();
    duckMics();
    return
  }

  if (mode == 'presentersAndAudience') {
    stopMonitor();
    unduckMics();
    return
  }

  if (mode == "autoDuck") {
    unduckMics(true);
    startMonitor();
  }

}

function processActions({ Type, Value, WidgetId }) {
  if (Type != 'released') return          // Ignore none released events
  if (WidgetId != config.panelId) return  // Ignore events from other widgets
  if (mode == Value) return                // Ignore events where the use selected the same value
  mode = Value
  applyMode();
}

async function checkInCall() {
  const mtrCall = await xapi.Status.MicrosoftTeams.Calling.InCall.get()
  const call = await xapi.Status.Call.get();
  return call?.[0]?.Status == 'Connected' || mtrCall == 'True'
}

function createMicLevels(samples) {
  const result = {}
  for (const key of micNames) {
    result[key] = new Array(samples).fill(0);
  }
  return result
}

function processNewCall(callType) {
  mode = config.defaultMode;
  const modeName = config.modeNames[mode];
  const buttonName = config.button.name;
  console.log(callType, 'Call Connected - Setting Mode:', config.defaultMode);
  xapi.Command.UserInterface.Extensions.Widget.SetValue({ WidgetId: config.panelId, Value: mode })
  applyMode();
  alert(`New Call Detected<br>Setting Audio Mode To: ${modeName}<br>Tap On [${buttonName}] Button To Select Other Modes.`);
}

function processCallEnd(callType) {
  console.log(callType, 'Call Ended - Stopping Monitor');
  stopMonitor();
}


function processAudioEvents(event) {

  audioEventCount = ++audioEventCount;
  clearTimeout(audioEventTimeout);
  audioEventTimeout = null;

  // console.log(event)

  const newLevels = flattenObject(event)

  //console.log('newLevels:', newLevels)

  for (const [micName, levels] of Object.entries(micLevels)) {
    micLevels[micName].shift();
    micLevels[micName].push(newLevels?.[micName] ?? levels[levels.length - 1]);
  }


  //console.debug('Levels:', micLevels)

  let aboveHighThreshold = false;
  let aboveLowThreshold = false;

  const averages = {}

  for (const [micName, levels] of Object.entries(micLevels)) {
    const sum = levels.reduce((partialSum, a) => partialSum + a, 0)
    const average = sum / levels.length;
    averages[micName] = average;
    micLevelAverages[micName].shift();
    micLevelAverages[micName].push(average);
    aboveHighThreshold = aboveHighThreshold ? aboveHighThreshold : average > config.threshold.high
    aboveLowThreshold = aboveLowThreshold ? aboveLowThreshold : average > config.threshold.low
  }

  if (aboveHighThreshold) {
    if (unduckTimer) console.error("Audio Levels Above Average")
    clearTimeout(unduckTimer)
    unduckTimer = null
    duckMics();
  }

  if (!aboveLowThreshold && !unduckTimer) {
    console.warn("Audio Levels Below Average")
    unduckTimer = setTimeout(() => {
      if (config.debug) console.log('Unducking Timeout Triggered')
      unduckMics();
    }, config.unduck.timeout * 1000)
  }

  if (audioEventCount == averageLogFequency) {
    if (config.debug) console.log('Audio Level Averages:\n' + JSON.stringify(micLevelAverages))
    audioEventCount = 0;
  }

  startAudioEventTimeout();

}

async function checkGainLevel() {
  const inputs = await xapi.Config.Audio.Input.get();
  const { Ethernet, Microphone, USBMicrophone } = inputs;
  if (Ethernet) return typeof Ethernet?.[0]?.Channel?.[0].Gain != 'undefined' ? 'Gain' : 'Level'
  if (Microphone) return Microphone.some(mic => typeof mic?.Gain != 'undefined') ? 'Gain' : 'Level'
  if (USBMicrophone) return typeof USBMicrophone?.[0]?.Gain != 'undefined' ? 'Gain' : 'Level'
}

function flattenObject(obj) {
  let result = {};
  for (const i in obj) {
    if (!obj.hasOwnProperty(i)) continue;
    if ((typeof obj[i]) == 'object' && obj[i] !== null) {
      const flatObject = flattenObject(obj[i], obj?.id);
      for (const x in flatObject) {
        if (!flatObject.hasOwnProperty(x)) continue;
        const id = obj[i]?.id ?? i;
        const key = (x == 'VuMeter') ? id : ((i == 'SubId') ? x : id + '.' + x)
        result[key] = flatObject[x];
      }
    } else {
      if (i != 'VuMeter') continue
      result[i] = parseInt(obj[i]);
    }
  }
  return result;
}


function startMonitor() {
  listener = xapi.Event.Audio.Input.Connectors.on(processAudioEvents);
  const monitoringMicNames = startVuMeterConnectors.map(({ ConnectorType, ConnectorId }) => ConnectorType + '.' + ConnectorId);
  console.log('Starting Audio Monitor:', ...monitoringMicNames);

  micLevelAverages = createMicLevels(averageLogFequency)
  micLevels = createMicLevels(config.samples)

  startVuMeterConnectors.forEach(({ ConnectorId, ConnectorType }) => {
    if (ConnectorId) {
      xapi.Command.Audio.VuMeter.Start({ ConnectorId, ConnectorType, IntervalMs: sampleInterval, Source: "BeforeAEC" });
    } else {
      xapi.Command.Audio.VuMeter.Start({ ConnectorType, IntervalMs: sampleInterval, Source: "BeforeAEC" });
    }
  })

  startAudioEventTimeout();
}

function stopMonitor() {
  console.log('Stopping Audio Monitor');
  clearTimeout(audioEventTimeout);
  if (listener) {
    listener();
    listener = () => void 0;
  }
  xapi.Command.Audio.VuMeter.StopAll();
}

function duckMics() {
  if (ducked) return
  console.log('Ducking Mics:', duck);
  const level = config.levels.duck;
  config.duck.forEach(mic => setInputLevelGain({ level, ...mic }))
  ducked = true;
}

function unduckMics(forceUnduck = false) {
  if (!ducked && !forceUnduck) return
  console.log('Unducking Mics:', duck);
  const level = config.levels.unduck;
  config.duck.forEach(mic => setInputLevelGain({ level, ...mic }))
  ducked = false;
}


function startAudioEventTimeout() {
  clearTimeout(audioEventTimeout);
  const timeout = 2 * sampleInterval;
  if (config.debug) console.debug('Starting Audio Event Timemout - delay:', timeout, ' ms');
  audioEventTimeout = setTimeout(() => {
    if (config.debug) console.debug('Audio Event Timeout Reached - Triggering Audio Events Process')
    processAudioEvents({})
  }, timeout);
}

async function setInputLevelGain({ ConnectorType, ConnectorId, SubId, level }) {
  const supportedTypes = ['Ethernet', 'Microphone', 'USBInterface', 'USBMicrophone']
  if (!supportedTypes.includes(ConnectorType)) {
    throw new Error(`Unsupported Audio Input Type [${ConnectorType}]`)
  }
  const mic = `${ConnectorType}.${ConnectorId}${SubId ? '.' + SubId : ''}`

  if (SubId) {
    await xapi.Config.Audio.Input[ConnectorType][ConnectorId].Channel[SubId][gainLevel].set(level);
    if (config.debug) console.log(`Mic: ${mic} - ${gainLevel}: ${level}`);
  } else {
    await xapi.Config.Audio.Input[ConnectorType][ConnectorId][gainLevel].set(level);
    if (config.debug) console.log(`Mic: ${mic} - ${gainLevel}: ${level}`);
  }
}


function consolidateConnectors(inputArray) {
  const uniqueConnectors = new Map(); // Use a Map to store unique combinations

  inputArray.forEach(item => {
    // Create a unique key based on ConnectorType and ConnectorId
    const key = `${item.ConnectorType}-${item.ConnectorId}`;

    // If this combination hasn't been added yet, add it to the Map
    if (!uniqueConnectors.has(key)) {
      uniqueConnectors.set(key, {
        ConnectorType: item.ConnectorType,
        ConnectorId: item.ConnectorId
      });
    }
  });

  // Convert the Map values back into an array
  return Array.from(uniqueConnectors.values());
}

function createMicStrings(inputArray) {

  const uniqueConnectors = new Map(); // Use a Map to store unique combinations

  inputArray.forEach(({ ConnectorType, ConnectorId, SubId }) => {
    // Create a unique key based on ConnectorType and ConnectorId
    const key = `${ConnectorType}-${ConnectorId}-${SubId ?? ''}`;

    // If this combination hasn't been added yet, add it to the Map
    if (!uniqueConnectors.has(key)) {
      uniqueConnectors.set(key, {
        ConnectorType,
        ConnectorId,
        SubId
      });
    }
  });

  const namesArray = Array.from(uniqueConnectors.values());

  return namesArray.map(({ ConnectorType, ConnectorId, SubId }) => ConnectorType + '.' + ConnectorId + (SubId ? '.' + SubId : ''));

}

function alert(Text = "", Duration = 10) {
  if (!config.showAlerts) return
  console.log('Displaying Alert:', Text)
  xapi.Command.UserInterface.Message.Alert.Display(
    { Duration, Target: "Controller", Text, Title: config.button.name });
}


async function createPanel() {
  const { icon, color, name, location } = config.button;
  const panelId = config.panelId;

  const order = await panelOrder(panelId);

  const values = Object.keys(config.modeNames).map(mode => {
    return `<Value><Key>${mode}</Key><Name>${config.modeNames[mode].replace(/&/g, "&amp;")}</Name></Value>`
  });


  const mtrDevice = await xapi.Command.MicrosoftTeams.List({ Show: 'Installed' })
    .then(() => true)
    .catch(() => false)


  const panelLocation = mtrDevice ? (location == 'Hidden' ? location : 'ControlPanel') : location;

  const panel = `
    <Extensions>
      <Panel>
        <Origin>local</Origin>
        <Location>${panelLocation}</Location>
        <Icon>${icon}</Icon>
        <Color>${color}</Color>
        <Name>${name}</Name>
        ${order}
        <ActivityType>Custom</ActivityType>
        <Page>
          <Name>${name}</Name>
          <Row>
            <Widget>
              <WidgetId>${panelId}</WidgetId>
              <Type>GroupButton</Type>
              <Options>size=4;columns=1</Options>
              <ValueSpace>
                ${values}
              </ValueSpace>
            </Widget>
          </Row>
          <Options>hideRowNames=1</Options>
        </Page>
      </Panel>
    </Extensions>`;

  return xapi.Command.UserInterface.Extensions.Panel.Save({ PanelId: panelId }, panel);
}


async function panelOrder(panelId) {
  const list = await xapi.Command.UserInterface.Extensions.List({ ActivityType: "Custom" });
  const panels = list?.Extensions?.Panel
  if (!panels) return ''
  const existingPanel = panels.find(panel => panel.PanelId == panelId)
  if (!existingPanel) return ''
  return `<Order>${existingPanel.Order}</Order>`
}