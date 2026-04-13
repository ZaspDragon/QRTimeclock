const SHEET_NAMES = {
  employees: 'ApprovedEmployees',
  punches: 'TimeLogs',
  managers: 'ManagerUsers',
  timesheets: 'Timesheets'
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const action = data.action;

    let result;
    switch (action) {
      case 'getEmployee':
        result = getEmployee_(data.employeeCode);
        break;
      case 'recordPunch':
        result = recordPunch_(data.employeeCode, data.action);
        break;
      case 'managerLogin':
        result = managerLogin_(data.username, data.password);
        break;
      case 'getTimesheets':
        result = getTimesheets_(data.workDate);
        break;
      case 'signTimesheet':
        result = signTimesheet_(data.timesheetId, data.managerName);
        break;
      default:
        result = { ok: false, message: 'Unknown action.' };
    }

    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ ok: false, message: err.message || String(err) });
  }
}

function doGet() {
  return jsonOut_({ ok: true, message: 'QR Time Clock backend is running.' });
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const employees = getOrCreateSheet_(ss, SHEET_NAMES.employees, [
    'EmployeeCode', 'EmployeeName', 'Active'
  ]);

  const punches = getOrCreateSheet_(ss, SHEET_NAMES.punches, [
    'PunchId', 'TimestampISO', 'WorkDate', 'EmployeeCode', 'EmployeeName', 'Action', 'LocalTime'
  ]);

  const managers = getOrCreateSheet_(ss, SHEET_NAMES.managers, [
    'Username', 'Password', 'ManagerName', 'Active'
  ]);

  const timesheets = getOrCreateSheet_(ss, SHEET_NAMES.timesheets, [
    'TimesheetId', 'WorkDate', 'EmployeeCode', 'EmployeeName', 'ClockIn', 'LunchStart', 'LunchEnd', 'ClockOut', 'TotalHours', 'ManagerSigned', 'SignedAtISO'
  ]);

  if (employees.getLastRow() === 1) {
    employees.getRange(2, 1, 3, 3).setValues([
      ['EMP001', 'John Doe', 'TRUE'],
      ['EMP002', 'Jane Smith', 'TRUE'],
      ['EMP003', 'Marcus Reed', 'TRUE']
    ]);
  }

  if (managers.getLastRow() === 1) {
    managers.getRange(2, 1, 1, 4).setValues([
      ['manager', 'password123', 'Warehouse Manager', 'TRUE']
    ]);
  }

  [employees, punches, managers, timesheets].forEach(sheet => {
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });
}

function getEmployee_(employeeCode) {
  const row = findByFirstCol_(SHEET_NAMES.employees, employeeCode);
  if (!row || String(row.Active).toUpperCase() !== 'TRUE') {
    return { ok: false, message: 'Employee code not found or inactive.' };
  }

  return {
    ok: true,
    employee: {
      code: row.EmployeeCode,
      name: row.EmployeeName
    }
  };
}

function recordPunch_(employeeCode, action) {
  const employee = getEmployee_(employeeCode);
  if (!employee.ok) return employee;

  const tz = Session.getScriptTimeZone() || 'America/New_York';
  const now = new Date();
  const workDate = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const localTime = Utilities.formatDate(now, tz, 'hh:mm:ss a');
  const punchId = Utilities.getUuid();

  const actionMap = {
    clockIn: 'CLOCK_IN',
    startLunch: 'LUNCH_START',
    endLunch: 'LUNCH_END',
    clockOut: 'CLOCK_OUT'
  };

  const normalizedAction = actionMap[action];
  if (!normalizedAction) return { ok: false, message: 'Bad punch action.' };

  const punchSheet = getSheet_(SHEET_NAMES.punches);
  punchSheet.appendRow([
    punchId,
    now.toISOString(),
    workDate,
    employee.employee.code,
    employee.employee.name,
    normalizedAction,
    localTime
  ]);

  upsertTimesheet_(workDate, employee.employee.code, employee.employee.name, normalizedAction, now, localTime, tz);

  return {
    ok: true,
    entry: {
      punchId,
      workDate,
      employeeCode: employee.employee.code,
      employeeName: employee.employee.name,
      action: normalizedAction,
      localTime
    }
  };
}

function managerLogin_(username, password) {
  const sheet = getSheet_(SHEET_NAMES.managers);
  const rows = rowsToObjects_(sheet);
  const match = rows.find(r => String(r.Username).trim() === String(username).trim() && String(r.Password).trim() === String(password).trim() && String(r.Active).toUpperCase() === 'TRUE');

  if (!match) return { ok: false, message: 'Manager username or password is incorrect.' };

  return {
    ok: true,
    manager: {
      username: match.Username,
      name: match.ManagerName || match.Username
    }
  };
}

function getTimesheets_(workDate) {
  const sheet = getSheet_(SHEET_NAMES.timesheets);
  let rows = rowsToObjects_(sheet);
  if (workDate) rows = rows.filter(r => String(r.WorkDate) === String(workDate));

  rows.sort((a, b) => {
    if (a.WorkDate === b.WorkDate) return String(a.EmployeeName).localeCompare(String(b.EmployeeName));
    return String(b.WorkDate).localeCompare(String(a.WorkDate));
  });

  return {
    ok: true,
    rows: rows.map(r => ({
      timesheetId: r.TimesheetId,
      workDate: r.WorkDate,
      employeeCode: r.EmployeeCode,
      employeeName: r.EmployeeName,
      clockIn: r.ClockIn,
      lunchStart: r.LunchStart,
      lunchEnd: r.LunchEnd,
      clockOut: r.ClockOut,
      totalHours: r.TotalHours,
      managerSigned: r.ManagerSigned
    }))
  };
}

function signTimesheet_(timesheetId, managerName) {
  const sheet = getSheet_(SHEET_NAMES.timesheets);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idIndex = headers.indexOf('TimesheetId');
  const managerIndex = headers.indexOf('ManagerSigned');
  const signedAtIndex = headers.indexOf('SignedAtISO');
  const workDateIndex = headers.indexOf('WorkDate');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIndex]) === String(timesheetId)) {
      sheet.getRange(r + 1, managerIndex + 1).setValue(managerName);
      sheet.getRange(r + 1, signedAtIndex + 1).setValue(new Date().toISOString());
      return { ok: true, message: 'Timesheet signed.', workDate: values[r][workDateIndex] };
    }
  }

  return { ok: false, message: 'Timesheet not found.' };
}

function upsertTimesheet_(workDate, employeeCode, employeeName, action, actionDate, localTime, tz) {
  const sheet = getSheet_(SHEET_NAMES.timesheets);
  const rows = rowsToObjects_(sheet);
  let existing = rows.find(r => String(r.WorkDate) === String(workDate) && String(r.EmployeeCode) === String(employeeCode));

  if (!existing) {
    existing = {
      TimesheetId: Utilities.getUuid(),
      WorkDate: workDate,
      EmployeeCode: employeeCode,
      EmployeeName: employeeName,
      ClockIn: '',
      LunchStart: '',
      LunchEnd: '',
      ClockOut: '',
      TotalHours: '',
      ManagerSigned: '',
      SignedAtISO: ''
    };
    sheet.appendRow([
      existing.TimesheetId,
      existing.WorkDate,
      existing.EmployeeCode,
      existing.EmployeeName,
      existing.ClockIn,
      existing.LunchStart,
      existing.LunchEnd,
      existing.ClockOut,
      existing.TotalHours,
      existing.ManagerSigned,
      existing.SignedAtISO
    ]);
    rows.push(existing);
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rowIndex = values.findIndex((row, i) => i > 0 && String(row[0]) === String(existing.TimesheetId));
  if (rowIndex === -1) throw new Error('Timesheet row missing.');

  const rowNumber = rowIndex + 1;
  const col = name => headers.indexOf(name) + 1;

  if (action === 'CLOCK_IN') sheet.getRange(rowNumber, col('ClockIn')).setValue(localTime);
  if (action === 'LUNCH_START') sheet.getRange(rowNumber, col('LunchStart')).setValue(localTime);
  if (action === 'LUNCH_END') sheet.getRange(rowNumber, col('LunchEnd')).setValue(localTime);
  if (action === 'CLOCK_OUT') sheet.getRange(rowNumber, col('ClockOut')).setValue(localTime);

  const updated = rowToObject_(headers, sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0]);
  const totalHours = calculateHours_(updated, workDate, tz);
  sheet.getRange(rowNumber, col('TotalHours')).setValue(totalHours);
}

function calculateHours_(row, workDate, tz) {
  if (!row.ClockIn || !row.ClockOut) return '';

  const clockIn = parseLocalDateTime_(workDate, row.ClockIn, tz);
  const clockOut = parseLocalDateTime_(workDate, row.ClockOut, tz);
  if (!clockIn || !clockOut) return '';

  let minutes = (clockOut.getTime() - clockIn.getTime()) / 60000;

  if (row.LunchStart && row.LunchEnd) {
    const lunchStart = parseLocalDateTime_(workDate, row.LunchStart, tz);
    const lunchEnd = parseLocalDateTime_(workDate, row.LunchEnd, tz);
    if (lunchStart && lunchEnd) {
      minutes -= (lunchEnd.getTime() - lunchStart.getTime()) / 60000;
    }
  }

  const hours = Math.max(0, minutes / 60);
  return hours.toFixed(2);
}

function parseLocalDateTime_(workDate, localTime, tz) {
  try {
    const combined = workDate + ' ' + localTime;
    return new Date(Utilities.formatDate(new Date(combined), tz, "yyyy-MM-dd'T'HH:mm:ss"));
  } catch (err) {
    return null;
  }
}

function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name);
  return sheet;
}

function rowsToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).map(row => rowToObject_(headers, row));
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

function findByFirstCol_(sheetName, firstColValue) {
  const rows = rowsToObjects_(getSheet_(sheetName));
  return rows.find(r => String(Object.values(r)[0]).trim() === String(firstColValue).trim()) || null;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
