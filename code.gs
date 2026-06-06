const SHEET_MEDIA = 'Media';
const SHEET_USERS = 'Users';

// ฟังก์ชันสร้างฐานข้อมูลอัตโนมัติ
function initDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ตรวจสอบและสร้าง Sheet: Media
  let mediaSheet = ss.getSheetByName(SHEET_MEDIA);
  if (!mediaSheet) {
    mediaSheet = ss.insertSheet(SHEET_MEDIA);
    // เพิ่ม Grade และ Indicator ในหัวตาราง
    mediaSheet.appendRow(['ID', 'Title', 'Description', 'Category', 'Thumbnail', 'Link', 'CreatedDate', 'Status', 'Grade', 'Indicator']);
    mediaSheet.appendRow(['123456789', 'ตัวอย่างสื่อการสอน', 'คำอธิบายตัวอย่าง', 'ภาษาไทย', 'https://images.unsplash.com/photo-1509228468518-180dd4864904?ixlib=rb-1.2.1&auto=format&fit=crop&w=500&q=60', '#', new Date().toISOString(), 'Active', 'ป.1', 'ท 1.1 ป.1/1']);
  }
  
  // 2. ตรวจสอบและสร้าง Sheet: Users
  let usersSheet = ss.getSheetByName(SHEET_USERS);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(SHEET_USERS);
    usersSheet.appendRow(['Username', 'Password', 'Role']);
    usersSheet.appendRow(['admin', 'admin123', 'admin']);
  }
}

function doGet(e) {
  initDB();
  const action = e.parameter.action;
  if (action === 'getMedia') {
    return createJsonResponse(getMediaData());
  }
  return createJsonResponse({ status: 'error', message: 'Invalid Action' });
}

function doPost(e) {
  initDB();
  
  if (typeof e === 'undefined' || typeof e.postData === 'undefined') {
    return createJsonResponse({ status: 'error', message: 'No POST data' });
  }

  let requestData;
  try {
    requestData = JSON.parse(e.postData.contents);
  } catch (error) {
    return createJsonResponse({ status: 'error', message: 'Invalid JSON' });
  }

  const action = e.parameter.action;
  
  if (action === 'login') {
    return createJsonResponse(checkLogin(requestData.username, requestData.password));
  } else if (action === 'addMedia') {
    return createJsonResponse(addMediaData(requestData));
  } else if (action === 'updateMedia') {
    return createJsonResponse(updateMediaData(requestData));
  } else if (action === 'deleteMedia') {
    return createJsonResponse(deleteMediaData(requestData.ID));
  }

  return createJsonResponse({ status: 'error', message: 'Invalid Action' });
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------- ส่วนจัดการข้อมูล ----------------
function getMediaData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MEDIA);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const result = [];
  
  if(data.length <= 1) return { status: 'success', data: [] };
  
  for (let i = 1; i < data.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return { status: 'success', data: result.reverse() }; 
}

function addMediaData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MEDIA);
  const newRow = [
    data.ID, 
    data.Title, 
    data.Description, 
    data.Category, 
    data.Thumbnail, 
    data.Link, 
    data.CreatedDate, 
    data.Status,
    data.Grade || '',      // บันทึก ระดับชั้น (คอลัมน์ I)
    data.Indicator || ''   // บันทึก ตัวชี้วัด (คอลัมน์ J)
  ];
  sheet.appendRow(newRow);
  return { status: 'success', message: 'Added successfully' };
}

function updateMediaData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MEDIA);
  const values = sheet.getDataRange().getValues();
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] == data.ID) {
      sheet.getRange(i + 1, 2).setValue(data.Title);
      sheet.getRange(i + 1, 3).setValue(data.Description);
      sheet.getRange(i + 1, 4).setValue(data.Category);
      sheet.getRange(i + 1, 5).setValue(data.Thumbnail);
      sheet.getRange(i + 1, 6).setValue(data.Link);
      sheet.getRange(i + 1, 8).setValue(data.Status);
      sheet.getRange(i + 1, 9).setValue(data.Grade || '');      // อัปเดต ระดับชั้น (คอลัมน์ I)
      sheet.getRange(i + 1, 10).setValue(data.Indicator || '');  // อัปเดต ตัวชี้วัด (คอลัมน์ J)
      return { status: 'success', message: 'Updated successfully' };
    }
  }
  return { status: 'error', message: 'ID not found' };
}

function deleteMediaData(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MEDIA);
  const values = sheet.getDataRange().getValues();
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { status: 'success', message: 'Deleted successfully' };
    }
  }
  return { status: 'error', message: 'ID not found' };
}

function checkLogin(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const values = sheet.getDataRange().getValues();
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] == username && values[i][1] == password) {
      return { status: 'success', role: values[i][2] };
    }
  }
  return { status: 'error', message: 'Invalid Credentials' };
}
