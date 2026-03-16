/**
 * TussScore - Google Apps Script Backend
 * ระบบค้นหาผลสอบออนไลน์ โรงเรียนเตรียมอุดมศึกษาภาคใต้
 *
 * วิธีใช้:
 * 1. สร้าง Google Sheet ใหม่
 * 2. ไปที่ Extensions > Apps Script
 * 3. วางโค้ดนี้ใน Code.gs
 * 4. Deploy > New deployment > Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. คัดลอก URL ที่ได้ไปใส่ในระบบ
 */

const SHEET_NAME = 'TussScore_Data';

const HEADERS = [
  '__backendId',
  'user_id', 'username', 'password',
  'exam_id', 'exam_name',
  'student_exam_number', 'student_id_number', 'student_name',
  'school', 'study_plan',
  'math_score', 'thai_score', 'english_score', 'science_score', 'social_score',
  'total_score', 'rank',
  'created_at',
  'config_type', 'search_field',
  'show_school', 'show_plan', 'show_rank', 'show_avg'
];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    // ตั้งค่า header row
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f48fb1');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    // สร้าง admin default
    createDefaultAdminRecord(sheet);
  }

  return sheet;
}

function createDefaultAdminRecord(sheet) {
  const adminId = generateId();
  const row = HEADERS.map(h => {
    const map = {
      '__backendId': adminId,
      'user_id': 'user_001',
      'username': 'admin',
      'password': 'tusscore',
      'created_at': new Date().toISOString()
    };
    return map[h] !== undefined ? map[h] : '';
  });
  sheet.appendRow(row);
}

function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2, 9);
}

function doGet(e) {
  try {
    const action = e.parameter.action || 'getAll';

    if (action === 'getAll') {
      const data = getAllData();
      return jsonResponse({ isOk: true, data: data });
    }

    return jsonResponse({ isOk: false, error: 'Unknown action' });
  } catch (error) {
    return jsonResponse({ isOk: false, error: error.toString() });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const data = body.data;
    let result;

    if (action === 'create') {
      result = createRecord(data);
    } else if (action === 'update') {
      result = updateRecord(data);
    } else if (action === 'delete') {
      result = deleteRecord(data.__backendId);
    } else {
      result = { isOk: false, error: 'Unknown action: ' + action };
    }

    return jsonResponse(result);
  } catch (error) {
    return jsonResponse({ isOk: false, error: error.toString() });
  }
}

function jsonResponse(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function getAllData() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  return values
    .map(row => {
      const obj = {};
      HEADERS.forEach((header, i) => {
        // แปลง number ให้ถูกต้อง
        const val = row[i];
        if (typeof val === 'number') {
          obj[header] = val;
        } else if (val === '' || val === null || val === undefined) {
          obj[header] = '';
        } else {
          obj[header] = String(val);
        }
      });
      return obj;
    })
    .filter(row => {
      // กรองแถวว่าง
      return row.__backendId !== '';
    });
}

function createRecord(data) {
  try {
    const sheet = getSheet();
    const newId = generateId();

    const row = HEADERS.map(header => {
      if (header === '__backendId') return newId;
      const val = data[header];
      return (val !== undefined && val !== null) ? val : '';
    });

    sheet.appendRow(row);
    return { isOk: true, data: { ...data, __backendId: newId } };
  } catch (error) {
    return { isOk: false, error: { message: error.toString() } };
  }
}

function updateRecord(data) {
  try {
    const sheet = getSheet();
    const backendId = data.__backendId;
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) return { isOk: false, error: { message: 'Record not found' } };

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let rowIndex = -1;

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(backendId)) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex === -1) return { isOk: false, error: { message: 'Record not found: ' + backendId } };

    const row = HEADERS.map(header => {
      if (header === '__backendId') return backendId;
      const val = data[header];
      return (val !== undefined && val !== null) ? val : '';
    });

    sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([row]);
    return { isOk: true, data };
  } catch (error) {
    return { isOk: false, error: { message: error.toString() } };
  }
}

function deleteRecord(backendId) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) return { isOk: false, error: { message: 'Record not found' } };

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let rowIndex = -1;

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(backendId)) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex === -1) return { isOk: false, error: { message: 'Record not found' } };

    sheet.deleteRow(rowIndex);
    return { isOk: true };
  } catch (error) {
    return { isOk: false, error: { message: error.toString() } };
  }
}
