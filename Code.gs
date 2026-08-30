/**
 * ===== CAIQC Head Sign — Apps Script Backend =====
 * ระบบ Preview + ลงนามอิเล็กทรอนิกส์สำหรับหัวหน้าแผนกปฏิบัติการกลาง
 * อ้างอิงข้อมูลจากชีต DB_CAIQC โดยตรง (ไม่พึ่ง Report/Forreport sheet)
 */

// ---------- CONFIG (แก้ตรงนี้ให้ตรงกับของจริง) ----------
const SPREADSHEET_ID = '1Qcffb8TgzvG7b48kzyjAOSc_sLhnX465O_gT7GTcUck';
const DB_SHEET_NAME = 'DB_CAIQC';
const AUDIT_SHEET_NAME = 'AuditLog';

// อีเมลของหัวหน้าที่มีสิทธิ์ลงนาม — ใส่ได้หลายคน (ตรงกับบัญชี Google จริงของแต่ละคน)
const AUTHORIZED_HEAD_EMAILS = [
  'jeeraphong.sriphat@gmail.com', // TODO: แก้เป็นอีเมลจริงของหัวหน้าคนที่ 1
  'lab34629@gmail.com',// 'head2@example.com', // เพิ่มบรรทัดแบบนี้สำหรับหัวหน้าคนถัดไป
];

// ชื่อคอลัมน์ที่ใช้ (ต้องตรงกับหัวตารางใน DB_CAIQC แถวที่ 1 เป๊ะๆ)
const COL = {
  ID: 'No.',
  RECORDER: 'ผู้บันทึก',
  DATE: 'วันที่ (พ.ศ.)',
  PLACE: 'สถานที่',
  TEST: 'Test',
  INSTRUMENT: 'Instrument',
  PROBLEM: 'สภาพปัญหา',
  SUMMARY: 'สรุปการแก้ไข',
  SIGNED: 'Signed_Head',
  SIGNED_BY: 'Signed_Head_By',
  SIGNED_DATETIME: 'Signed_Head_DateTime',
  SIGNED_HASH: 'Signed_Head_Hash',
  CAIQC_LINK: 'CAIQC link'
};

// ---------- CONFIG: การบันทึกลายเซ็นไว้ใช้ซ้ำ (แยกตามผู้ลงนามแต่ละคน) ----------
const SIGNATURE_PROPERTY_PREFIX = 'HEAD_SIGNATURE_IMG_'; // key จริงจะต่อท้ายด้วยอีเมลผู้ลงนาม
const REPORT_SHEET_NAME = 'Report';
const REPORT_SELECTOR_SHEET_NAME = 'ForReport'; // ชีตที่มี cell ตัวขับสูตร VLOOKUP จริง (ไม่ใช่ Report เอง)
const REPORT_SELECTOR_CELL = 'B2';   // cell ที่ใช้เลือกเลขที่บันทึก ในชีต ForReport ด้านบน
const REPORT_SIGNATURE_CELL = 'J41'; // cell สำหรับแทรกรูปลายเซ็นอิเล็กทรอนิกส์ (ไม่มีข้อความ/วันที่ปน)
const REPORT_DATE_CELL = 'AA41';     // cell สำหรับวันที่ลงนาม (ว/ด/ปปปป พ.ศ.) — anchor ของ merged range AA41:AD41
const REPORT_PDF_FOLDER_ID = '1FS8hzJ-qQODHFOpg6F3l-FdF_agiumAG';     // ใส่ Drive Folder ID ถ้าต้องการเก็บ PDF ไว้ในโฟลเดอร์เฉพาะ (ว่าง = เก็บที่ My Drive root)

// ---------- ENTRY POINT ----------
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ลงนามอนุมัติ CAIQC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ใช้เพื่อ include ไฟล์ HTML/CSS/JS แยกไฟล์ (ถ้ามีในอนาคต)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------- AUTH ----------
function checkAuthorized_() {
  const email = Session.getActiveUser().getEmail();
  if (!AUTHORIZED_HEAD_EMAILS.includes(email)) {
    throw new Error('บัญชีนี้ไม่มีสิทธิ์ลงนามอนุมัติ (' + (email || 'ไม่พบอีเมล') + ')');
  }
  return email;
}

// เรียกจาก client ตอนโหลดหน้าเว็บ เพื่อเช็คสิทธิ์ก่อนแสดง UI
function getCurrentUserStatus() {
  const email = Session.getActiveUser().getEmail();
  return {
    email: email || '',
    authorized: AUTHORIZED_HEAD_EMAILS.includes(email)
  };
}

// ---------- SHEET HELPERS ----------
function getSheet_(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('ไม่พบชีตชื่อ "' + name + '"');
  return sheet;
}

function getHeaderIndexMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i; });
  return map;
}

// ---------- READ: รายการที่ยังไม่ได้ลงนาม ----------
function getPendingRecords() {
  const sheet = getSheet_(DB_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headerMap = getHeaderIndexMap_(sheet);

  const idxSigned = headerMap[COL.SIGNED];
  const idxId = headerMap[COL.ID];
  const idxDate = headerMap[COL.DATE];
  const idxPlace = headerMap[COL.PLACE];
  const idxTest = headerMap[COL.TEST];
  const idxProblem = headerMap[COL.PROBLEM];
  const idxRecorder = headerMap[COL.RECORDER];

  const results = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row[idxId]) continue; // แถวว่าง
    const signed = row[idxSigned] === true || row[idxSigned] === 'TRUE';
    if (signed) continue; // แสดงเฉพาะที่ยังไม่ได้เซ็น
    results.push({
      id: row[idxId],
      date: formatDate_(row[idxDate]),
      place: row[idxPlace],
      test: row[idxTest],
      problem: row[idxProblem],
      recorder: row[idxRecorder]
    });
  }
  return results;
}

// ---------- READ: รายละเอียดเต็มของ 1 record สำหรับ preview ----------
// หมายเหตุ: คืนค่าเป็น JSON string (ไม่ใช่ object ตรงๆ) เพื่อเลี่ยงบั๊กของ
// google.script.run ที่บางครั้งทำให้ object ที่ซับซ้อนถูกส่งมาเป็น null ที่ฝั่ง client
function getRecordDetail(caiqcNo) {
  const sheet = getSheet_(DB_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headerMap = getHeaderIndexMap_(sheet);
  const idxId = headerMap[COL.ID];

  for (let r = 1; r < data.length; r++) {
    if (data[r][idxId] === caiqcNo) {
      const obj = {};
      Object.keys(headerMap).forEach(colName => {
        obj[colName] = sanitizeValue_(data[r][headerMap[colName]]);
      });
      obj._rowNumber = r + 1; // เก็บเลขแถวจริงไว้ใช้ตอนเขียนกลับ
      return JSON.stringify(obj);
    }
  }
  throw new Error('ไม่พบเลขที่บันทึก ' + caiqcNo);
}

// แปลงค่าทุกชนิดให้เป็น string/boolean/number ธรรมดา ก่อนส่งผ่าน google.script.run
// (Date object บางครั้งทำให้ google.script.run คืนค่า null แบบเงียบๆ)
function sanitizeValue_(value) {
  if (value === null || value === undefined) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  if (typeof value === 'boolean' || typeof value === 'number') return value;
  return String(value);
}

// ---------- WRITE: บันทึกการลงนาม ----------
function signRecord(caiqcNo, expectedRowNumber, signatureDataUrl, rememberSignature) {
  const signerEmail = checkAuthorized_(); // โยน error ถ้าไม่มีสิทธิ์
  if (!signatureDataUrl) {
    throw new Error('กรุณาวาดลายเซ็นก่อนยืนยัน');
  }
  const sheet = getSheet_(DB_SHEET_NAME);
  const headerMap = getHeaderIndexMap_(sheet);
  const idxId = headerMap[COL.ID];
  const idxSigned = headerMap[COL.SIGNED];

  const rowValues = sheet.getRange(expectedRowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];

  // กันเซ็นผิดแถว/ข้อมูลเปลี่ยนไปตั้งแต่ preview จนถึงตอนกดยืนยัน
  if (rowValues[idxId] !== caiqcNo) {
    throw new Error('ข้อมูลไม่ตรงกับที่ preview ไว้ กรุณาโหลดหน้าใหม่แล้วลองอีกครั้ง');
  }
  if (rowValues[idxSigned] === true || rowValues[idxSigned] === 'TRUE') {
    throw new Error('รายการนี้ถูกลงนามไปแล้ว');
  }

  const now = new Date();
  const hash = computeRowHash_(rowValues, headerMap);

  sheet.getRange(expectedRowNumber, headerMap[COL.SIGNED] + 1).setValue(true);
  sheet.getRange(expectedRowNumber, headerMap[COL.SIGNED_BY] + 1).setValue(signerEmail);
  sheet.getRange(expectedRowNumber, headerMap[COL.SIGNED_DATETIME] + 1).setValue(now);
  sheet.getRange(expectedRowNumber, headerMap[COL.SIGNED_HASH] + 1).setValue(hash);

  appendAuditLog_(caiqcNo, 'SIGN_HEAD', signerEmail, 'ลงนามอนุมัติผ่านหน้าเว็บ');

  // ---- สร้าง PDF จาก Report sheet (ไม่ทำให้การลงนามล้มเหลว ถ้าขั้นตอนนี้พลาด) ----
  let pdfUrl = '';
  try {
    pdfUrl = generateSignedReportPdf_(caiqcNo, signerEmail, now, signatureDataUrl);
    const idxLink = headerMap[COL.CAIQC_LINK];
    if (idxLink !== undefined && pdfUrl) {
      sheet.getRange(expectedRowNumber, idxLink + 1).setValue(pdfUrl);
    }
    appendAuditLog_(caiqcNo, 'PDF_EXPORTED', signerEmail, pdfUrl);
  } catch (pdfErr) {
    appendAuditLog_(caiqcNo, 'PDF_EXPORT_FAILED', signerEmail, pdfErr.message);
  }

  // ---- ลบลายเซ็นที่บันทึกไว้ของผู้ลงนามคนนี้ทิ้งเสมอ หลังลงนามสำเร็จ ----
  // (ครั้งถัดไปจะเริ่มจาก canvas ว่างเปล่าให้วาดใหม่ทุกครั้ง ไม่ต้องกดล้างเอง)
  PropertiesService.getScriptProperties().deleteProperty(SIGNATURE_PROPERTY_PREFIX + signerEmail);

  return { success: true, signedBy: signerEmail, signedAt: now.toISOString(), pdfUrl: pdfUrl };
}

// ---------- ลายเซ็นที่บันทึกไว้ใช้ซ้ำ (ของบัญชีที่ login อยู่ตอนนี้เท่านั้น) ----------
function getSavedSignature() {
  const email = checkAuthorized_();
  return PropertiesService.getScriptProperties().getProperty(SIGNATURE_PROPERTY_PREFIX + email) || '';
}

function clearSavedSignature() {
  const email = checkAuthorized_();
  PropertiesService.getScriptProperties().deleteProperty(SIGNATURE_PROPERTY_PREFIX + email);
  return { success: true };
}

// ---------- READ: รายการที่เซ็นไปแล้ว (สำหรับหน้ายกเลิกการลงนาม) ----------
function getSignedRecords() {
  const sheet = getSheet_(DB_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headerMap = getHeaderIndexMap_(sheet);

  const idxSigned = headerMap[COL.SIGNED];
  const idxId = headerMap[COL.ID];
  const idxTest = headerMap[COL.TEST];
  const idxProblem = headerMap[COL.PROBLEM];
  const idxSignedBy = headerMap[COL.SIGNED_BY];
  const idxSignedDT = headerMap[COL.SIGNED_DATETIME];
  const idxLink = headerMap[COL.CAIQC_LINK];

  const results = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row[idxId]) continue;
    const signed = row[idxSigned] === true || row[idxSigned] === 'TRUE';
    if (!signed) continue;
    results.push({
      id: row[idxId],
      test: row[idxTest],
      problem: row[idxProblem],
      signedBy: row[idxSignedBy],
      signedDateTime: sanitizeValue_(row[idxSignedDT]),
      pdfUrl: idxLink !== undefined ? row[idxLink] : '',
      _rowNumber: r + 1
    });
  }
  return results;
}

// ---------- WRITE: ยกเลิกการลงนาม (ให้กลับไปเซ็นใหม่ได้) ----------
function revokeSignature(caiqcNo, expectedRowNumber, reason) {
  const actorEmail = checkAuthorized_(); // เฉพาะบัญชีหัวหน้าเท่านั้นที่ยกเลิกได้
  if (!reason || !reason.toString().trim()) {
    throw new Error('กรุณาระบุเหตุผลในการยกเลิกการลงนาม');
  }

  const sheet = getSheet_(DB_SHEET_NAME);
  const headerMap = getHeaderIndexMap_(sheet);
  const idxId = headerMap[COL.ID];
  const idxSigned = headerMap[COL.SIGNED];

  const rowValues = sheet.getRange(expectedRowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (rowValues[idxId] !== caiqcNo) {
    throw new Error('ข้อมูลไม่ตรงกับที่เลือกไว้ กรุณาโหลดหน้าใหม่แล้วลองอีกครั้ง');
  }
  const isSigned = rowValues[idxSigned] === true || rowValues[idxSigned] === 'TRUE';
  if (!isSigned) {
    throw new Error('รายการนี้ยังไม่ได้ลงนาม จึงไม่สามารถยกเลิกได้');
  }

  // เก็บข้อมูลการลงนามเดิมไว้ใน AuditLog ก่อนล้างค่า (ไม่ให้ประวัติหายไปไหน)
  const oldSignedBy = rowValues[headerMap[COL.SIGNED_BY]];
  const oldSignedDateTime = sanitizeValue_(rowValues[headerMap[COL.SIGNED_DATETIME]]);
  const oldHash = rowValues[headerMap[COL.SIGNED_HASH]];
  const idxLink = headerMap[COL.CAIQC_LINK];
  const oldPdfLink = idxLink !== undefined ? rowValues[idxLink] : '';

  appendAuditLog_(caiqcNo, 'REVOKE_SIGN_HEAD', actorEmail,
    'เหตุผล: ' + reason +
    ' | ลงนามเดิมโดย: ' + oldSignedBy +
    ' เมื่อ: ' + oldSignedDateTime +
    ' | hash เดิม: ' + oldHash +
    (oldPdfLink ? ' | PDF เดิม: ' + oldPdfLink : ''));

  // ล้างสถานะลงนามใน DB_CAIQC เพื่อให้กลับมาเซ็นใหม่ได้ (ไม่แตะคอลัมน์ CAIQC link เพื่อให้ยังเปิด PDF เดิมได้จนกว่าจะเซ็นใหม่)
  sheet.getRange(expectedRowNumber, headerMap[COL.SIGNED] + 1).setValue(false);
  sheet.getRange(expectedRowNumber, headerMap[COL.SIGNED_BY] + 1).setValue('');
  sheet.getRange(expectedRowNumber, headerMap[COL.SIGNED_DATETIME] + 1).setValue('');
  sheet.getRange(expectedRowNumber, headerMap[COL.SIGNED_HASH] + 1).setValue('');

  return { success: true, revokedBy: actorEmail };
}

// ---------- PDF: เขียนข้อมูลลง Report sheet แล้ว export เฉพาะชีตนั้นเป็น PDF ----------
function generateSignedReportPdf_(caiqcNo, signerEmail, signedAtDate, signatureDataUrl) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const reportSheet = ss.getSheetByName(REPORT_SHEET_NAME);
  if (!reportSheet) throw new Error('ไม่พบชีตชื่อ "' + REPORT_SHEET_NAME + '"');
  const selectorSheet = ss.getSheetByName(REPORT_SELECTOR_SHEET_NAME);
  if (!selectorSheet) throw new Error('ไม่พบชีตชื่อ "' + REPORT_SELECTOR_SHEET_NAME + '"');

  // 1. ตั้งเลขที่บันทึกใน cell ตัวเลือกที่ชีต ForReport (ตัวขับสูตร VLOOKUP จริง ไม่ใช่ Report เอง)
  selectorSheet.getRange(REPORT_SELECTOR_CELL).setValue(caiqcNo);
  SpreadsheetApp.flush();

  // 2. เขียนวันที่ลงนาม (ว/ด/ปปปป พ.ศ.) ลง AA41 ของ Report
  reportSheet.getRange(REPORT_DATE_CELL).setValue(toThaiDateString_(signedAtDate));

  // 3. แทรกรูปลายเซ็นอิเล็กทรอนิกส์ที่วาดจริงลง K41 ของ Report (ไม่มีข้อความปน)
  insertSignatureImage_(reportSheet, signatureDataUrl);
  SpreadsheetApp.flush();

  // 4. Export เฉพาะ Report sheet เป็น PDF
  const gid = reportSheet.getSheetId();
  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID +
    '/export?format=pdf&gid=' + gid +
    '&size=A4&portrait=true&fitw=true&gridlines=false' +
    '&printtitle=false&sheetnames=false&pagenumbers=false' +
    '&horizontal_alignment=CENTER&vertical_alignment=TOP';

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + token }
  });
  const safeFileName = caiqcNo.replace(/\s+/g, '_') + '_signed.pdf';
  const pdfBlob = response.getBlob().setName(safeFileName);

  // 5. บันทึกลง Drive
  const folder = REPORT_PDF_FOLDER_ID ? DriveApp.getFolderById(REPORT_PDF_FOLDER_ID) : null;
  const file = folder ? folder.createFile(pdfBlob) : DriveApp.createFile(pdfBlob);

  // ให้ผู้ที่มีลิงก์เปิดดูได้ — ปรับตามนโยบายความปลอดภัยขององค์กรได้ตามต้องการ
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

// ---------- แทรกรูปลายเซ็น (data URL จาก canvas) ลง cell ที่กำหนด ----------
function insertSignatureImage_(sheet, dataUrl) {
  const match = String(dataUrl).match(/^data:(image\/\w+);base64,(.+)$/);
  if (!match) throw new Error('รูปแบบข้อมูลลายเซ็นไม่ถูกต้อง');
  const contentType = match[1];
  const base64Data = match[2];
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, contentType, 'signature.png');

  removeExistingSignatureImages_(sheet);

  const targetRange = sheet.getRange(REPORT_SIGNATURE_CELL);
  targetRange.clearContent(); // ล้างข้อความเก่า (เช่นจากการทดสอบเวอร์ชันก่อนหน้า) ให้เหลือแค่รูปเท่านั้น

  const image = sheet.insertImage(blob, targetRange.getColumn(), targetRange.getRow());
  image.setWidth(200);
  image.setHeight(50);
}

// ลบรูปลายเซ็นเก่าที่เคยแทรกไว้ที่ cell เดิม (กันภาพซ้อนกันถ้ามีการรันซ้ำ)
function removeExistingSignatureImages_(sheet) {
  const targetRange = sheet.getRange(REPORT_SIGNATURE_CELL);
  const targetRow = targetRange.getRow();
  const targetCol = targetRange.getColumn();
  sheet.getImages().forEach(function(img) {
    const anchor = img.getAnchorCell();
    if (anchor.getRow() === targetRow && anchor.getColumn() === targetCol) {
      img.remove();
    }
  });
}

// ---------- UTIL ----------
function computeRowHash_(rowValues, headerMap) {
  // hash ของข้อมูลทั้งแถว ณ ขณะเซ็น ไว้ตรวจสอบย้อนหลังว่าเนื้อหาไม่ถูกแก้หลังเซ็น
  const plain = JSON.stringify(rowValues);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, plain, Utilities.Charset.UTF_8);
  return digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function appendAuditLog_(caiqcNo, action, actor, detail) {
  const sheet = getSheet_(AUDIT_SHEET_NAME);
  sheet.appendRow([new Date(), caiqcNo, action, actor, detail]);
}

function formatDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return value; // เผื่อเป็น string อยู่แล้ว
}

// แปลงเป็นวันที่ไทย ว/ด/ปปปป (พ.ศ.) เช่น 30/08/2569
function toThaiDateString_(date) {
  const dayMonth = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM');
  const gregorianYear = Number(Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy'));
  return dayMonth + '/' + (gregorianYear + 543);
}
