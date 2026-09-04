const FINE_PER_DAY = 5; // ค่าปรับวันละ 5 บาท
const MAX_BORROW_LIMIT = 5; // ยืมได้สูงสุด 5 เล่ม

// === การตั้งค่าชื่อชีตและคอลัมน์ (ให้ตรงกับใน Google Sheet ของคุณ) ===
const SHEET_BOOKS = "Books";
const SHEET_USERS = "Users";
const SHEET_RECORDS = "BorrowRecords";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    if (action === "borrow") {
      return handleBorrow(data.userId, data.bookId);
    } else if (action === "return") {
      return handleReturn(data.userId, data.bookId);
    } else {
      return responseJson({ error: "Invalid action" }, 400);
    }
  } catch (error) {
    return responseJson({ error: error.toString() }, 500);
  }
}

// ==========================================
// ฟังก์ชันสำหรับการยืมหนังสือ
// ==========================================
function handleBorrow(userId, bookId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(SHEET_USERS);
  const bookSheet = ss.getSheetByName(SHEET_BOOKS);
  const recordSheet = ss.getSheetByName(SHEET_RECORDS);

  // 1. ตรวจสอบผู้ใช้
  const userRow = findRow(userSheet, 1, userId); // สมมติว่า Column A (1) คือ UserID
  if (userRow === -1) {
    return responseJson({ error: "ไม่พบรหัสสมาชิกนี้ในระบบ" }, 404);
  }
  const userName = userSheet.getRange(userRow, 2).getValue();

  // 2. ตรวจสอบเงื่อนไขการยืม (จำนวนเล่มค้างส่ง และหนังสือเกินกำหนด)
  const records = recordSheet.getDataRange().getValues();
  let pendingCount = 0;
  let hasOverdue = false;
  const now = new Date();

  for (let i = 1; i < records.length; i++) { // ข้าม header (row 0)
    const recUserId = records[i][1]; // Column B (Index 1) คือ UserID
    const recStatus = records[i][6]; // Column G (Index 6) คือ Status
    const recDueDate = new Date(records[i][4]); // Column E (Index 4) คือ DueDate

    if (recUserId == userId && recStatus === 'pending') {
      pendingCount++;
      if (now > recDueDate) {
        hasOverdue = true;
      }
    }
  }

  if (hasOverdue) {
    return responseJson({ error: "ไม่สามารถยืมได้ เนื่องจากมีหนังสือค้างส่งเกินกำหนด!" }, 400);
  }
  if (pendingCount >= MAX_BORROW_LIMIT) {
    return responseJson({ error: `ไม่สามารถยืมได้ เนื่องจากยืมครบกำหนด ${MAX_BORROW_LIMIT} เล่มแล้ว` }, 400);
  }

  // 3. ตรวจสอบหนังสือ
  const bookRow = findRow(bookSheet, 1, bookId); // สมมติ Column A (1) คือ BookID
  if (bookRow === -1) {
    return responseJson({ error: "ไม่พบรหัสหนังสือเล่มนี้ในระบบ" }, 404);
  }
  
  const bookTitle = bookSheet.getRange(bookRow, 2).getValue();
  const bookStatus = bookSheet.getRange(bookRow, 3).getValue();
  
  if (bookStatus !== 'available') {
    return responseJson({ error: `หนังสือ "${bookTitle}" ไม่พร้อมให้บริการ (ถูกยืมไปแล้ว)` }, 400);
  }

  // 4. ดำเนินการยืม
  const borrowDate = new Date();
  const dueDate = new Date();
  dueDate.setDate(borrowDate.getDate() + 7); // กำหนดคืน 7 วัน
  const recordId = Utilities.getUuid();

  // บันทึกลง BorrowRecords (สมมติคอลัมน์เรียงตาม: RecordID, UserID, BookID, BorrowDate, DueDate, ReturnDate, Status, Fine)
  recordSheet.appendRow([
    recordId, 
    userId, 
    bookId, 
    borrowDate, 
    dueDate, 
    "", 
    "pending", 
    0
  ]);

  // อัปเดตสถานะหนังสือเป็น borrowed
  bookSheet.getRange(bookRow, 3).setValue('borrowed');

  return responseJson({ 
    success: true, 
    message: `ยืมหนังสือ "${bookTitle}" สำเร็จโดยคุณ ${userName}. กำหนดส่งคืนวันที่ ${Utilities.formatDate(dueDate, "GMT+7", "dd/MM/yyyy")}` 
  }, 200);
}

// ==========================================
// ฟังก์ชันสำหรับการคืนหนังสือ
// ==========================================
function handleReturn(userId, bookId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName(SHEET_RECORDS);
  const bookSheet = ss.getSheetByName(SHEET_BOOKS);

  // 1. ค้นหารายการยืมที่ยังไม่ได้คืน
  const records = recordSheet.getDataRange().getValues();
  let targetRow = -1;
  let dueDateStr = "";
  
  for (let i = 1; i < records.length; i++) {
    if (records[i][1] == userId && records[i][2] == bookId && records[i][6] === 'pending') {
      targetRow = i + 1; // Google Sheet row เริ่มที่ 1 แต่ array เริ่ม 0
      dueDateStr = records[i][4];
      break;
    }
  }

  if (targetRow === -1) {
    return responseJson({ error: "ไม่พบรายการยืมหนังสือเล่มนี้ หรือถูกคืนไปแล้ว" }, 400);
  }

  // 2. คำนวณค่าปรับ
  const dueDate = new Date(dueDateStr);
  const returnDate = new Date();
  let fineAmount = 0;
  
  if (returnDate > dueDate) {
    const timeDiff = returnDate.getTime() - dueDate.getTime();
    const daysLate = Math.ceil(timeDiff / (1000 * 3600 * 24));
    fineAmount = daysLate * FINE_PER_DAY;
  }

  // 3. อัปเดตตาราง BorrowRecords
  recordSheet.getRange(targetRow, 6).setValue(returnDate); // คอลัมน์ F (ReturnDate)
  recordSheet.getRange(targetRow, 7).setValue('returned'); // คอลัมน์ G (Status)
  recordSheet.getRange(targetRow, 8).setValue(fineAmount); // คอลัมน์ H (Fine)

  // 4. อัปเดตสถานะหนังสือให้ว่าง
  const bookRow = findRow(bookSheet, 1, bookId);
  let bookTitle = "Unknown";
  if (bookRow !== -1) {
    bookSheet.getRange(bookRow, 3).setValue('available');
    bookTitle = bookSheet.getRange(bookRow, 2).getValue();
  }

  let message = `คืนหนังสือ "${bookTitle}" สำเร็จ`;
  if (fineAmount > 0) {
    message += ` (เกินกำหนด! มีค่าปรับ ${fineAmount} บาท)`;
  }

  return responseJson({ 
    success: true, 
    message: message,
    fine: fineAmount
  }, 200);
}

// ==========================================
// Helper Functions
// ==========================================
function findRow(sheet, columnNumber, searchString) {
  const columnValues = sheet.getRange(1, columnNumber, sheet.getLastRow(), 1).getValues();
  for (let i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0].toString() === searchString.toString()) {
      return i + 1;
    }
  }
  return -1;
}

function responseJson(obj, statusCode) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// สำหรับ Test ผ่าน Browser เบื้องต้น
function doGet(e) {
  return responseJson({ message: "Library Google Apps Script API is running." }, 200);
}
