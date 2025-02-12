// ตั้งค่าตัวแปร
const SHEET_URL = "https://docs.google.com/spreadsheets/d/1qIhI8XVEr2zxzsLZvcAo_ksudevG_V3TWNPTA7fM0J8/edit?usp=sharing";
const HR_EMAIL = "finkmata@gmail.com";

function onFormSubmit(e) {
  const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName("Form Responses");
  const { values, range } = e;
  const row = range.getRow();
   if (new Date(startDate) > new Date(endDate)) {
    throw new Error('วันที่เริ่มต้นต้องไม่เกินวันที่สิ้นสุด');
  }

  const validLeaveTypes = ['ลาป่วย', 'ลาพักร้อน', 'ลากิจ'];
  if (!validLeaveTypes.includes(leaveType)) {
    throw new Error('ประเภทการลาไม่ถูกต้อง');
  }

  // ข้อมูลจาก Google Form
  const [timestamp, name, email, leaveType, startDate, endDate, reason] = values;

  // คำนวณจำนวนวันลา
  const days = calculateLeaveDays(startDate, endDate);

  // บันทึกจำนวนวันลาและสถานะ
  sheet.getRange(row, 8).setValue(days); // คอลัมน์ H
  sheet.getRange(row, 9).setValue("รออนุมัติ"); // คอลัมน์ I


  // ส่งอีเมลแจ้งหัวหน้า
  sendApprovalEmail(name, email, leaveType, startDate, endDate, days, reason);
  
}


// คำนวณจำนวนวันลา
function calculateLeaveDays(start, end) {
  const startDate = new Date(start);
  const endDate = new Date(end);
  return (endDate - startDate) / (1000 * 60 * 60 * 24) + 1;
}

// ส่งอีเมลอนุมัติการลา
function sendApprovalEmail(name, email, leaveType, start, end, days, reason) {
  const subject = `คำขอลา ${leaveType} จาก ${name}`;
  const webAppUrl = ScriptApp.getService().getUrl();
  const htmlBody = `
    <p>ชื่อ: ${name}</p>
    <p>ประเภทการลา: ${leaveType}</p>
    <p>วันที่: ${start} ถึง ${end} (รวม ${days} วัน)</p>
    <p>เหตุผล: ${reason}</p>
    <p>กรุณาดำเนินการผ่าน <a href="${webAppUrl}">ระบบจัดการการลา</a></p>
  `;
  GmailApp.sendEmail(HR_EMAIL, subject, '', { htmlBody: htmlBody });
}

// อนุมัติการลา
function approveLeave(row) {
  const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName("Form Responses");
  sheet.getRange(row, 9).setValue("อนุมัติแล้ว");
  updateLeaveBalance(row);
}
// ไม่อนุมัติการลา + บันทึกเหตุผล
function rejectLeave(row, reason) {
  try {
    const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName("Form Responses");
    
    // อัปเดตสถานะและเหตุผล
    sheet.getRange(row, 9).setValue("ไม่อนุมัติ"); // คอลัมน์ I
    sheet.getRange(row, 10).setValue(reason); // คอลัมน์ J
    
    // ส่งอีเมลแจ้งพนักงาน
    const data = sheet.getRange(row, 1, 1, 9).getValues()[0];
    const [timestamp, name, email] = data;
    
    const subject = `⛔ คำขอลาของคุณไม่ได้รับอนุมัติ | ${name}`;
    const htmlBody = `
      <p>เรียน ${name}</p>
      <p>คำขอลาของคุณระหว่างวันที่ ${data[4]} ถึง ${data[5]} <strong>ไม่ได้รับอนุมัติ</strong></p>
      <p><strong>เหตุผล:</strong> ${reason}</p>
      <p>สอบถามเพิ่มเติม: ${HR_EMAIL}</p>
      <hr>
      <small>ระบบจัดการการลาอัตโนมัติ</small>
    `;
    
    GmailApp.sendEmail(email, subject, '', { htmlBody: htmlBody });
    
    return true;
  } catch (error) {
    console.error(error);
    throw new Error('เกิดข้อผิดพลาดในการไม่อนุมัติการลา');
  }
}

// ดึงข้อมูลคำขอลา (แก้ไขให้ดึงสถานะทั้งหมด)
// แก้ไขฟังก์ชัน getLeaveRequests ให้เป็นเวอร์ชันเดียว
function getLeaveRequests() {
  const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName("Form Responses");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const requests = [];
  
  // กำหนดดัชนีคอลัมน์ด้วยชื่อ
  const COL = {
    TIMESTAMP: headers.indexOf('Timestamp'),
    NAME: headers.indexOf('ชื่อ'),
    EMAIL: headers.indexOf('อีเมล'),
    LEAVE_TYPE: headers.indexOf('ประเภทการลา'),
    START_DATE: headers.indexOf('วันที่เริ่มลา'),
    END_DATE: headers.indexOf('วันที่สิ้นสุด'),
    REASON: headers.indexOf('เหตุผล'),
    DAYS: headers.indexOf('จำนวนวันลา'),
    STATUS: headers.indexOf('สถานะ')
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[COL.STATUS] === "รออนุมัติ") {
      requests.push({
        row: i + 1,
        name: row[COL.NAME],
        leaveType: row[COL.LEAVE_TYPE],
        startDate: Utilities.formatDate(new Date(row[COL.START_DATE]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        endDate: Utilities.formatDate(new Date(row[COL.END_DATE]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        days: row[COL.DAYS],
        reason: row[COL.REASON]
      });
    }
  }
  return requests;
}

// อัปเดตจำนวนวันลาคงเหลือ
function updateLeaveBalance(row) {
  const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName("Form Responses");
  const data = sheet.getRange(row, 1, 1, 9).getValues()[0];
  const [name, leaveType, days] = [data[1], data[3], data[7]];

  const balanceSheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName("Leave Balance");
  const employeeData = balanceSheet.getDataRange().getValues();
  
  for (let i = 0; i < employeeData.length; i++) {
    if (employeeData[i][0] === name) {
      const currentBalance = employeeData[i][1];
      const newBalance = currentBalance - days;
      balanceSheet.getRange(i + 1, 2).setValue(newBalance);
      break;
    }
  }
}



function processLeaveRequest(formData) {
  try {
    const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName("Form Responses");

    // ตรวจสอบข้อมูลก่อนบันทึก
    if (!formData.name || !formData.email || !formData.startDate || !formData.endDate) {
      throw new Error("ข้อมูลไม่ครบถ้วน");
    }

    // บันทึกลง Google Sheets
    sheet.appendRow([
      new Date(),
      formData.name,
      formData.email,
      formData.leaveType,
      formData.startDate,
      formData.endDate,
      formData.reason || "ไม่มีเหตุผล",
      '',
      'รออนุมัติ'
    ]);
    
  } catch (error) {
    Logger.log("เกิดข้อผิดพลาด: " + error.message);
    throw error;
  }
}


// ดึงข้อมูลคำขอลา
function getLeaveRequests() {
  const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName("Form Responses");
  const data = sheet.getDataRange().getValues();
  const requests = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[8] === "รออนุมัติ") {
      requests.push({
        row: i + 1,
        name: row[1],
        leaveType: row[3],
        startDate: Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        endDate: Utilities.formatDate(new Date(row[5]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        days: row[7],
        reason: row[6]
      });
    }
  }
  return requests;
}
// แสดงหน้าเว็บ
function getLeaveStatus() {
  const sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName("Form Responses");
  const data = sheet.getDataRange().getValues();
  const userEmail = Session.getActiveUser().getEmail(); // ✅ ดึงอีเมลของผู้ใช้ที่ล็อกอิน
  const results = [];

  const headers = data[0]; // หัวข้อคอลัมน์
  const COL = {
    TIMESTAMP: headers.indexOf("Timestamp"),
    LEAVE_TYPE: headers.indexOf("ประเภทการลา"),
    START_DATE: headers.indexOf("วันที่เริ่มลา"),
    END_DATE: headers.indexOf("วันที่สิ้นสุด"),
    DAYS: headers.indexOf("จำนวนวันลา"),
    STATUS: headers.indexOf("สถานะ"),
    EMAIL: headers.indexOf("อีเมล")
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[COL.EMAIL] === userEmail) { // ✅ ดึงเฉพาะข้อมูลของผู้ใช้ที่ล็อกอินอยู่
      results.push({
        date: Utilities.formatDate(new Date(row[COL.TIMESTAMP]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        leaveType: row[COL.LEAVE_TYPE],
        startDate: row[COL.START_DATE],
        endDate: row[COL.END_DATE],
        days: row[COL.DAYS],
        status: row[COL.STATUS]
      });
    }
  }

  return results; // ✅ ส่งข้อมูลกลับไปที่ `status.html`
}

function doGet(e) {
  let page = e.parameter.page || "index"; // ตรวจสอบว่าต้องเปิดหน้าไหน
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle("ระบบจัดการการลา")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}
