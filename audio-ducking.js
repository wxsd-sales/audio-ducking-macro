/********************************************************
 * 
 * Author:              William Mills
 *                    	Solutions Engineer
 *                    	wimills@cisco.com
 *                    	Cisco Systems
 * 
 * 
 * Version: 1-0-0
 * Released: 09/11/25
 * 
 * Audio Ducking Macro
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
  button: {
    name: 'Audio Modes',
    color: '#f58142',
    icon: 'Sliders'
  },
  mics: [
    { ConnectorType: 'Microphone', ConnectorId: 1 },
    //{ ConnectorType: 'USBMicrophone', ConnectorId: 1 },
    //{ ConnectorType: 'Ethernet', ConnectorId: 1, SubId: 1 },
    //{ ConnectorType: 'Ethernet', ConnectorId: 2, SubId: 3 }
  ],
  duck: [
    //{ ConnectorType: 'Microphone', ConnectorId: 1 },
    { ConnectorType: 'Ethernet', ConnectorId: 1, SubId: 1 },
    { ConnectorType: 'Ethernet', ConnectorId: 2, SubId: 1 },
    { ConnectorType: 'Ethernet', ConnectorId: 3, SubId: 1 },
    { ConnectorType: 'Ethernet', ConnectorId: 4, SubId: 1 },
    { ConnectorType: 'Ethernet', ConnectorId: 5, SubId: 1 },
    { ConnectorType: 'Ethernet', ConnectorId: 6, SubId: 1 }
  ],
  threshold: {
    high: 30,
    low: 25
  },
  levels: {
    duck: 0,
    unduck: 30
  },
  unduck:{
    timeout: 5 // Unduck mics after 5 seconds
  },
  samples: 4,
  panelId: 'audioDucking'
}


/*********************************************************
 * Do not change below
**********************************************************/



const startVuMeterConnectors = consolidateConnectors(config.mics);
console.log('startVuMeterConnectors:', startVuMeterConnectors);

const micNames = createMicStrings(config.mics);
console.log('micNames:', micNames);

const duck = config.duck.map(({ ConnectorType, ConnectorId, SubId }) => ConnectorType + '.' + ConnectorId + (SubId ? '.' + SubId : ''));

let ducked = false;
let unduckTimer;
let listener;
let micLevels;
let mode;

init();

async function init(){

  micLevels = createMicLevels(config.samples)
  await createPanel();

  xapi.Event.UserInterface.Extensions.Widget.Action.on(processActions);

  xapi.Status.Call.on(({ ghost, Status }) => {
    if (Status && Status == 'Connected') return applyMode();
    if (ghost) return stopMonitor();
  })

  const widgets = await xapi.Status.UserInterface.Extensions.Widget.get()
  const selection = widgets.find(widget => widget.WidgetId == config.panelId);
  const value = selection?.Value;

  mode = value && value != '' ? value : 'presenters'


  applyMode();

}


async function applyMode(){
  console.log('Mode:', mode);
  const inCall = checkInCall();
  if(!inCall) return
  if(mode == 'presenters') {
    stopMonitor();
    duckMics();
    return
  }

  if(mode == 'audience') return unduckMics();
  if(mode == "presentersAndAudience") return startMonitor();

}

function processActions({Type, Value, WidgetId}){
  if(Type != 'released') return
  if(WidgetId != config.panelId) return
  mode = Value
  applyMode();
}

async function checkInCall() {
  const call = await xapi.Status.Call.get();
  return call?.[0]?.Status == 'Connected'
}

function createMicLevels(samples){
  const result = {}
  for (const key of micNames) {
    result[key] = new Array(samples).fill(0);
  }
  return result
}


function processAudioEvents(event) {

  const newLevels = flattenObject(event)

  for (const [micName, levels] of Object.entries(micLevels)) {
    micLevels[micName].shift();
    micLevels[micName].push(newLevels?.[micName] ?? levels[levels.length - 1]);
  }

  let aboveHighThreshold = false;
  let aboveLowThreshold = false;

  const averages = {}

  for (const [micName, levels] of Object.entries(micLevels)) {
    const sum = levels.reduce((partialSum, a) => partialSum + a, 0)
    const average = sum / levels.length;
    averages[micName] = average;
    aboveHighThreshold = aboveHighThreshold ? aboveHighThreshold : average > config.threshold.high
    aboveLowThreshold = aboveLowThreshold ? aboveLowThreshold : average > config.threshold.low
  }

  if (aboveHighThreshold) {
    if(unduckTimer) console.error("Above Average")
    clearTimeout(unduckTimer)
    unduckTimer = null
    duckMics();
  }

  if (!aboveLowThreshold && !unduckTimer) {
    console.warn("Below Average")
    unduckTimer = setTimeout(()=>{
      console.log('Unducking Timer')
      unduckMics();
      }, config.unduck.timeout * 1000)
  }

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
  startVuMeterConnectors.forEach(({ ConnectorId, ConnectorType }) => {
    if (ConnectorId) {
      xapi.Command.Audio.VuMeter.Start({ ConnectorId, ConnectorType, IntervalMs: 100, Source: "BeforeAEC" });
    } else {
      xapi.Command.Audio.VuMeter.Start({ ConnectorType, IntervalMs: 100, Source: "BeforeAEC" });
    }
  })
}

function stopMonitor() {
  console.log('Stopping Audio');
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

function unduckMics() {
  if (!ducked) return
  console.log('Unducking Mics:', duck);
  const level = config.levels.unduck;
  config.duck.forEach(mic => setInputLevelGain({ level, ...mic }))
  ducked = false;
}

async function setInputLevelGain({ ConnectorType, ConnectorId, SubId, level }) {
  const supportedTypes = ['Ethernet', 'Microphone', 'USBInterface', 'USBMicrophone']
  if (!supportedTypes.includes(ConnectorType)) {
    throw new Error(`Unsupported Audio Input Type [${ConnectorType}]`)
  }
  const mic = `${ConnectorType}.${ConnectorId}${SubId ? '.' + SubId : ''}`
  console.log(`Setting ${mic} - Level: [${level}]`)

  if (SubId) {
    try {
      await xapi.Config.Audio.Input[ConnectorType][ConnectorId].Channel[SubId].Gain.set(level);
      return console.log('Mic:', mic, '- Gain:', level, ' - Set')
    } catch (e) {

      try {
        await xapi.Config.Audio.Input[ConnectorType][ConnectorId].Channel[SubId].Level.set(level);
        return console.log('Mic:', mic, '- Gain:', level, ' - Set')
      } catch (e) {
        return console.warn('Error Setting Mic:', mic, '- Level:', level)
      }
    }

  } else {
    try {
      await xapi.Config.Audio.Input[ConnectorType][ConnectorId].Gain.set(level);
      return console.log('Mic:', mic, '- Level:', level, ' - Set')
    } catch (e) {
      try {
        await xapi.Config.Audio.Input[ConnectorType][ConnectorId].Level.set(level);
        return console.log('Mic:', mic, '- Level:', level, ' - Set')
      } catch (e) {
        return console.warn('Error Setting Mic:', mic, '- Level:', level)
      }
    }
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


async function createPanel(){
  const {icon, color, name}= config.button;
  const panelId = config.panelId;

  const order = await panelOrder(panelId)
  const panel = `
    <Extensions>
      <Panel>
        <Origin>local</Origin>
        <Location>HomeScreenAndCallControls</Location>
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
                <Value>
                  <Key>presenters</Key>
                  <Name>Presenters Only</Name>
                </Value>
                <Value>
                  <Key>audience</Key>
                  <Name>Audience Only</Name>
                </Value>
                <Value>
                  <Key>presentersAndAudience</Key>
                  <Name>Auto Adjust</Name>
                </Value>
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
