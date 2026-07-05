const DEFAULT_SPREADSHEET_ID = 'XXX';

function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const customId = props.getProperty('CUSTOM_SPREADSHEET_ID');
  const targetId = customId && customId.trim() !== '' ? customId.trim() : DEFAULT_SPREADSHEET_ID;

  try {
    return SpreadsheetApp.openById(targetId);
  } catch (e) {
    Logger.log('Could not open by ID, falling back to default: ' + e.message);
    return SpreadsheetApp.openById(DEFAULT_SPREADSHEET_ID);
  }
}

function getSpreadsheetId() {
  const props = PropertiesService.getScriptProperties();
  const customId = props.getProperty('CUSTOM_SPREADSHEET_ID');
  return customId && customId.trim() !== '' ? customId.trim() : DEFAULT_SPREADSHEET_ID;
}

function setSpreadsheetId(newId) {
  if (!newId || newId.trim() === '') return { success: false, message: 'Invalid ID' };
  PropertiesService.getScriptProperties().setProperty('CUSTOM_SPREADSHEET_ID', newId.trim());
  initDatabase();
  return { success: true, newId: newId.trim() };
}

function getFormattedTimestamp() {
  const now = new Date();
  return Utilities.formatDate(now, "Asia/Bangkok", "d/M/yyyy, H:mm:ss");
}

function doGet(e) {
  initDatabase();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('POCT QualiAgentics - ระบบควบคุมคุณภาพเครื่องตรวจ POCT')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Initialize Database tables and default settings if not exists
 */
function initDatabase() {
  const ss = getSpreadsheet();
  
  // Sheet 1: Settings
  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.appendRow(['Category', 'Value']);
    
    const defaultSettings = [
      ['TestType', 'Blood Glucose (DTX)'],
      ['TestType', 'Blood Lactate'],
      ['Department', 'สามัญหญิง'],
      ['Department', 'สามัญชาย'],
      ['Department', 'พิเศษ 5'],
      ['Department', 'เคมีบำบัด'],
      ['Department', 'มะเร็งนรีเวช'],
      ['Department', 'รังสีวินิจฉัย'],
      ['Department', 'รังสีรักษา'],
      ['Department', 'รังสีร่วมรักษา'],
      ['Department', 'ทันตกรรม'],
      ['Department', 'ผู้ป่วยนอกทั่วไป'],
      ['Department', 'วิสัญญีพยาบาล'],
      ['Department', 'ประกายแสง 5'],
      ['Department', 'เวชศาสตร์นิวเคลียร์'],
      ['Department', 'ส่องกล้อง'],
      ['Contact', 'พยาบาลประจำเคาน์เตอร์'],
      ['Contact', 'นักเทคนิคการแพทย์'],
      ['Supervisor', 'ปวรวรรชน์ ภูริคัมภีร์'],
      ['Supervisor', 'จักรพงษ์ ณ รังษี'],
      ['Status', 'พร้อมใช้'],
      ['Status', 'เลิกใช้'],
      ['Status', 'ระหว่างซ่อม']
    ];
    defaultSettings.forEach(r => settingsSheet.appendRow(r));
  }
  
  // Sheet 2: Devices
  let devicesSheet = ss.getSheetByName('Devices');
  if (!devicesSheet) {
    devicesSheet = ss.insertSheet('Devices');
    devicesSheet.appendRow([
      'DeviceID', 'DeviceName', 'TestType', 'Model', 'SerialNumber', 
      'Department', 'Contact', 'StartDate', 'Status', 'CreatedAt'
    ]);
    
    devicesSheet.appendRow([
      'DEV-001', 'Glucose Meter #1', 'Blood Glucose (DTX)', 'Accu-Chek Inform II', 'SN-9876541',
      'สามัญหญิง', 'พยาบาลประจำเคาน์เตอร์', '2025-01-10', 'พร้อมใช้', getFormattedTimestamp()
    ]);
    devicesSheet.appendRow([
      'DEV-002', 'Lactate Scout #1', 'Blood Lactate', 'Lactate Scout 4', 'SN-1234567',
      'วิสัญญีพยาบาล', 'นักเทคนิคการแพทย์', '2025-02-01', 'พร้อมใช้', getFormattedTimestamp()
    ]);
  }
  
  // Sheet 3: Reagents
  let reagentsSheet = ss.getSheetByName('Reagents');
  if (!reagentsSheet) {
    reagentsSheet = ss.insertSheet('Reagents');
    reagentsSheet.appendRow([
      'ReagentID', 'TestType', 'ReagentName', 'LotNumber', 'ExpiryDate', 'CreatedAt'
    ]);
    
    reagentsSheet.appendRow([
      'REA-001', 'Blood Glucose (DTX)', 'Accu-Chek Performa Strips', 'LOT-GLU-2026A', '2026-12-31', getFormattedTimestamp()
    ]);
    reagentsSheet.appendRow([
      'REA-002', 'Blood Lactate', 'Lactate Scout Test Strips', 'LOT-LAC-2026B', '2026-10-15', getFormattedTimestamp()
    ]);
  }
  
  // Sheet 4: IQC_Controls (Mean & SD ONLY)
  let controlsSheet = ss.getSheetByName('IQC_Controls');
  if (!controlsSheet) {
    controlsSheet = ss.insertSheet('IQC_Controls');
    controlsSheet.appendRow([
      'ControlID', 'TestType', 'ControlName', 'LotNumber', 'LevelsCount',
      'L1_Mean', 'L1_SD', 'L2_Mean', 'L2_SD', 'L3_Mean', 'L3_SD',
      'StartDate', 'ExpiryDate', 'CreatedAt'
    ]);
    
    controlsSheet.appendRow([
      'CTL-001', 'Blood Glucose (DTX)', 'Glucose Control Solution', 'CTL-GLU-99', 2,
      55, 3.5, 290, 12, '', '',
      '2026-01-01', '2026-12-31', getFormattedTimestamp()
    ]);
  }
  
  // Sheet 5: IQC_Results
  let resultsSheet = ss.getSheetByName('IQC_Results');
  if (!resultsSheet) {
    resultsSheet = ss.insertSheet('IQC_Results');
    resultsSheet.appendRow([
      'ResultID', 'Timestamp', 'TestDate', 'DeviceID', 'ReagentLot', 'ControlLot',
      'L1_Result', 'L2_Result', 'L3_Result', 'TesterName', 'SupervisorName',
      'Status', 'ViolatedRules', 'Notes'
    ]);
    
    resultsSheet.appendRow([
      'RES-001', getFormattedTimestamp(), '2026-06-01', 'DEV-001', 'LOT-GLU-2026A', 'CTL-GLU-99',
      '54', '288', '', 'พยาบาลประจำเคาน์เตอร์', 'ปวรวรรชน์ ภูริคัมภีร์',
      'In Control', 'None', 'ผลปกติ'
    ]);
  }
  
  // Sheet 6: EQA_Registry
  let eqaSheet = ss.getSheetByName('EQA_Registry');
  if (!eqaSheet) {
    eqaSheet = ss.insertSheet('EQA_Registry');
    eqaSheet.appendRow(['EQAID', 'TestType', 'Provider', 'SampleCode', 'TargetDate', 'Status']);
  }
}

/**
 * Fetch all initial application data from Google Sheet
 */
function getInitialData() {
  const ss = getSpreadsheet();
  
  const settingsSheet = ss.getSheetByName('Settings');
  const devicesSheet = ss.getSheetByName('Devices');
  const reagentsSheet = ss.getSheetByName('Reagents');
  const controlsSheet = ss.getSheetByName('IQC_Controls');
  const resultsSheet = ss.getSheetByName('IQC_Results');
  
  return {
    spreadsheetId: getSpreadsheetId(),
    settings: getSheetRows(settingsSheet),
    devices: getSheetRows(devicesSheet),
    reagents: getSheetRows(reagentsSheet),
    controls: getSheetRows(controlsSheet),
    results: getSheetRows(resultsSheet)
  };
}

function getSheetRows(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = row[j];
      if (val instanceof Date) {
        const headerName = headers[j];
        if (headerName === 'CreatedAt' || headerName === 'Timestamp') {
          val = Utilities.formatDate(val, "Asia/Bangkok", "d/M/yyyy, H:mm:ss");
        } else {
          val = Utilities.formatDate(val, "Asia/Bangkok", "yyyy-MM-dd");
        }
      }
      obj[headers[j]] = val;
    }
    rows.push(obj);
  }
  return rows;
}

/**
 * Add / Update / Delete Device
 */
function addDevice(deviceData) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Devices');
  const rows = sheet.getDataRange().getValues();
  
  const count = rows.length;
  const deviceId = 'DEV-' + String(count).padStart(3, '0');
  
  sheet.appendRow([
    deviceId,
    deviceData.deviceName,
    deviceData.testType,
    deviceData.model,
    deviceData.serialNumber,
    deviceData.department,
    deviceData.contact,
    deviceData.startDate,
    deviceData.status || 'พร้อมใช้',
    getFormattedTimestamp()
  ]);
  
  return { success: true, deviceId: deviceId };
}

function updateDevice(deviceId, deviceData) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Devices');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === deviceId) {
      const rowIdx = i + 1;
      sheet.getRange(rowIdx, 2).setValue(deviceData.deviceName);
      sheet.getRange(rowIdx, 3).setValue(deviceData.testType);
      sheet.getRange(rowIdx, 4).setValue(deviceData.model);
      sheet.getRange(rowIdx, 5).setValue(deviceData.serialNumber);
      sheet.getRange(rowIdx, 6).setValue(deviceData.department);
      sheet.getRange(rowIdx, 7).setValue(deviceData.contact);
      sheet.getRange(rowIdx, 8).setValue(deviceData.startDate);
      sheet.getRange(rowIdx, 9).setValue(deviceData.status);
      return { success: true };
    }
  }
  return { success: false, message: 'Not found' };
}

function updateDeviceStatus(deviceId, newStatus) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Devices');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === deviceId) {
      sheet.getRange(i + 1, 9).setValue(newStatus);
      return { success: true };
    }
  }
  return { success: false, message: 'Device not found' };
}

function deleteDevice(deviceId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Devices');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === deviceId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'Device not found' };
}

/**
 * Add / Update / Delete Reagent
 */
function addReagent(reagentData) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Reagents');
  const count = sheet.getDataRange().getValues().length;
  const reagentId = 'REA-' + String(count).padStart(3, '0');
  
  sheet.appendRow([
    reagentId,
    reagentData.testType,
    reagentData.reagentName,
    reagentData.lotNumber,
    reagentData.expiryDate,
    getFormattedTimestamp()
  ]);
  
  return { success: true, reagentId: reagentId };
}

function updateReagent(reagentId, reagentData) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Reagents');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reagentId) {
      const rowIdx = i + 1;
      sheet.getRange(rowIdx, 2).setValue(reagentData.testType);
      sheet.getRange(rowIdx, 3).setValue(reagentData.reagentName);
      sheet.getRange(rowIdx, 4).setValue(reagentData.lotNumber);
      sheet.getRange(rowIdx, 5).setValue(reagentData.expiryDate);
      return { success: true };
    }
  }
  return { success: false, message: 'Not found' };
}

function deleteReagent(reagentId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Reagents');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reagentId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'Reagent not found' };
}

/**
 * Add / Update / Delete IQC Control (Mean & SD ONLY)
 */
function addIQCControl(controlData) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('IQC_Controls');
  const count = sheet.getDataRange().getValues().length;
  const controlId = 'CTL-' + String(count).padStart(3, '0');
  
  sheet.appendRow([
    controlId,
    controlData.testType,
    controlData.controlName,
    controlData.lotNumber,
    controlData.levelsCount,
    controlData.l1Mean || '', controlData.l1SD || '',
    controlData.l2Mean || '', controlData.l2SD || '',
    controlData.l3Mean || '', controlData.l3SD || '',
    controlData.startDate,
    controlData.expiryDate,
    getFormattedTimestamp()
  ]);
  
  return { success: true, controlId: controlId };
}

function updateIQCControl(controlId, controlData) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('IQC_Controls');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === controlId) {
      const rowIdx = i + 1;
      sheet.getRange(rowIdx, 2).setValue(controlData.testType);
      sheet.getRange(rowIdx, 3).setValue(controlData.controlName);
      sheet.getRange(rowIdx, 4).setValue(controlData.lotNumber);
      sheet.getRange(rowIdx, 5).setValue(controlData.levelsCount);
      
      sheet.getRange(rowIdx, 6).setValue(controlData.l1Mean || '');
      sheet.getRange(rowIdx, 7).setValue(controlData.l1SD || '');

      sheet.getRange(rowIdx, 8).setValue(controlData.l2Mean || '');
      sheet.getRange(rowIdx, 9).setValue(controlData.l2SD || '');

      sheet.getRange(rowIdx, 10).setValue(controlData.l3Mean || '');
      sheet.getRange(rowIdx, 11).setValue(controlData.l3SD || '');

      sheet.getRange(rowIdx, 12).setValue(controlData.startDate);
      sheet.getRange(rowIdx, 13).setValue(controlData.expiryDate);
      return { success: true };
    }
  }
  return { success: false, message: 'Not found' };
}

function deleteIQCControl(controlId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('IQC_Controls');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === controlId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'Control not found' };
}

/**
 * Add / Update / Delete IQC Result
 */
function addIQCResult(resultData) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('IQC_Results');
  const count = sheet.getDataRange().getValues().length;
  const resultId = 'RES-' + String(count).padStart(3, '0');
  
  sheet.appendRow([
    resultId,
    getFormattedTimestamp(),
    resultData.testDate,
    resultData.deviceId,
    resultData.reagentLot,
    resultData.controlLot,
    resultData.l1Result || '',
    resultData.l2Result || '',
    resultData.l3Result || '',
    resultData.testerName,
    resultData.supervisorName,
    resultData.status,
    resultData.violatedRules,
    resultData.notes || ''
  ]);
  
  return { success: true, resultId: resultId };
}

function updateIQCResult(resultId, resultData) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('IQC_Results');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === resultId) {
      const rowIdx = i + 1;
      sheet.getRange(rowIdx, 3).setValue(resultData.testDate);
      sheet.getRange(rowIdx, 4).setValue(resultData.deviceId);
      sheet.getRange(rowIdx, 5).setValue(resultData.reagentLot);
      sheet.getRange(rowIdx, 6).setValue(resultData.controlLot);
      sheet.getRange(rowIdx, 7).setValue(resultData.l1Result || '');
      sheet.getRange(rowIdx, 8).setValue(resultData.l2Result || '');
      sheet.getRange(rowIdx, 9).setValue(resultData.l3Result || '');
      sheet.getRange(rowIdx, 10).setValue(resultData.testerName);
      sheet.getRange(rowIdx, 11).setValue(resultData.supervisorName);
      sheet.getRange(rowIdx, 12).setValue(resultData.status);
      sheet.getRange(rowIdx, 13).setValue(resultData.violatedRules);
      sheet.getRange(rowIdx, 14).setValue(resultData.notes || '');
      return { success: true };
    }
  }
  return { success: false, message: 'Result not found' };
}

function deleteIQCResult(resultId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('IQC_Results');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === resultId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'Result not found' };
}

function addSettingOption(category, value) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Settings');
  sheet.appendRow([category, value]);
  return { success: true };
}
