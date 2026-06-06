const SHEET_MEDIA = 'Media';
const SHEET_USERS = 'Users';

// ฟังก์ชันสร้างฐานข้อมูลอัตโนมัติ
function initDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ตรวจสอบและสร้าง Sheet: Media
  let mediaSheet = ss.getSheetByName(SHEET_MEDIA);
  if (!mediaSheet) {
    mediaSheet = ss.insertSheet(SHEET_MEDIA);
    // สร้างหัวตาราง
    mediaSheet.appendRow(['ID', 'Title', 'Description', 'Category', 'Thumbnail', 'Link', 'CreatedDate', 'Status']);
    // ใส่ข้อมูลตัวอย่าง 1 แถวเพื่อให้ไม่ว่างเปล่า
    mediaSheet.appendRow(['123456789', 'ตัวอย่างสื่อการสอน', 'คำอธิบายตัวอย่าง', 'ทั่วไป', 'https://images.unsplash.com/photo-1509228468518-180dd4864904?ixlib=rb-1.2.1&auto=format&fit=crop&w=500&q=60', '#', new Date().toISOString(), 'Active']);
  }
  
  // 2. ตรวจสอบและสร้าง Sheet: Users (สำหรับล็อกอิน)
  let usersSheet = ss.getSheetByName(SHEET_USERS);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(SHEET_USERS);
    // สร้างหัวตาราง
    usersSheet.appendRow(['Username', 'Password', 'Role']);
    // สร้าง Admin เริ่มต้น (รหัสผ่านคือ admin123)
    usersSheet.appendRow(['admin', 'admin123', 'admin']);
  }
}

function doGet(e) {
  // รันระบบสร้างฐานข้อมูลอัตโนมัติ
  initDB();
  
  const action = e.parameter.action;
  if (action === 'getMedia') {
    return createJsonResponse(getMediaData());
  }
  return createJsonResponse({ status: 'error', message: 'Invalid Action' });
}

function doPost(e) {
  // รันระบบสร้างฐานข้อมูลอัตโนมัติ
  initDB();
  
  // ป้องกัน Error กรณีไม่มีข้อมูลส่งมา (Preflight check)
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

// ฟังก์ชันส่งคืนค่ากลับไปยังเว็บไซต์
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
  
  // ถ้ามีแต่หัวตาราง ส่ง Array ว่างกลับไป
  if(data.length <= 1) return { status: 'success', data: [] };
  
  for (let i = 1; i < data.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return { status: 'success', data: result.reverse() }; // สลับเอาอันใหม่ขึ้นก่อน
}

function addMediaData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MEDIA);
  const newRow = [
    data.ID, data.Title, data.Description, data.Category, 
    data.Thumbnail, data.Link, data.CreatedDate, data.Status
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

---

### ⚠️ ขั้นตอนแก้ปัญหา "Failed to fetch" แบบชัวร์ 100%

หลังจากวางโค้ดใน Apps Script แล้ว **วิธีการ Deploy สำคัญมากครับ** ให้ทำตามนี้เป๊ะๆ:

1. กดเมนู **การทำให้ใช้งานได้ (Deploy)** > **การทำให้ใช้งานได้รายการใหม่ (New deployment)**
2. คลิกที่รูปเฟือง ⚙️ ทางซ้ายมือ เลือก **เว็บแอป (Web app)**
3. ในช่อง **แอปทำงานในฐานะ (Execute as)** ให้เลือก 👉 **ฉัน (Me)** *(อีเมลของคุณ)*
4. ในช่อง **ผู้มีสิทธิ์เข้าถึง (Who has access)** ให้เลือก 👉 **ทุกคน (Anyone)** *(สำคัญมาก! ถ้าไม่เลือก Everyone จะขึ้น Failed to fetch ทันที)*
5. กดปุ่ม **การทำให้ใช้งานได้ (Deploy)**
6. คัดลอก **URL เว็บแอป** ที่ลงท้ายด้วย `/exec` อันใหม่ล่าสุดไปใส่ทับในตัวแปร `const SCRIPT_URL = '...';` ในไฟล์ `index.html` ของคุณ

เพียงเท่านี้ หน้าเว็บจะเชื่อมต่อฐานข้อมูลได้โดยไม่เกิด Error และจะสร้างตาราง Excel ให้อัตโนมัติเลยครับ ลองทดสอบดูได้เลยครับ!
