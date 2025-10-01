function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('ระบบออกเลขหนังสือ');
}

function saveData(data) {
  const ss = SpreadsheetApp.openById('1LHV8jSSsQ7TAc1Emi_888jliem003l2ys4EIEqMh_GQ');
  const bookSheet = ss.getSheetByName('book_id');
  const sentSheet = ss.getSheetByName('book_sent');
  const lockSheet = ss.getSheetByName('lock') || ss.insertSheet('lock');

  if (!bookSheet) throw new Error('Sheet book_id not found');
  if (!sentSheet) throw new Error('Sheet book_sent not found');

  // ระบบ locking เพื่อป้องกัน race condition
  const maxRetries = 5;
  const lockTimeout = 10000; // 10 วินาที
  let lockAcquired = false;
  let retries = 0;

  while (!lockAcquired && retries < maxRetries) {
    try {
      // ตรวจสอบ lock ที่มีอยู่
      const currentTime = new Date().getTime();
      const lockData = lockSheet.getRange(1, 1, 1, 2).getValues()[0];
      const existingLock = lockData[0];
      const lockTime = lockData[1] ? new Date(lockData[1]).getTime() : 0;

      // ถ้ามี lock และยังไม่หมดอายุ ให้รอ
      if (existingLock && (currentTime - lockTime < lockTimeout)) {
        retries++;
        if (retries >= maxRetries) {
          return { success: false, message: 'ระบบกำลังทำงาน กรุณาลองใหม่อีกครั้งในอีกสักครู่' };
        }
        Utilities.sleep(200 + Math.random() * 300); // รอแบบสุ่ม 200-500ms
        continue;
      }

      // พยายามได้ lock
      const lockId = Utilities.getUuid();
      lockSheet.getRange(1, 1).setValue(lockId);
      lockSheet.getRange(1, 2).setValue(new Date());
      
      // รอสักนิดแล้วตรวจสอบว่าได้ lock จริงหรือไม่
      Utilities.sleep(100);
      const verifyLock = lockSheet.getRange(1, 1).getValue();
      
      if (verifyLock === lockId) {
        lockAcquired = true;
      } else {
        retries++;
        if (retries >= maxRetries) {
          return { success: false, message: 'ระบบกำลังทำงาน กรุณาลองใหม่อีกครั้ง' };
        }
        Utilities.sleep(200 + Math.random() * 300);
        continue;
      }

    } catch (error) {
      retries++;
      if (retries >= maxRetries) {
        return { success: false, message: 'เกิดข้อผิดพลาด กรุณาลองใหม่อีกครั้ง' };
      }
      Utilities.sleep(200 + Math.random() * 300);
      continue;
    }
  }

  if (!lockAcquired) {
    return { success: false, message: 'ไม่สามารถประมวลผลได้ในขณะนี้ กรุณาลองใหม่อีกครั้ง' };
  }

  try {
    const timestamp = new Date();

    // Format timestamp in Thai date and time
    const thaiMonths = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
    const thaiYear = timestamp.getFullYear() + 543;
    const thaiDate = `${timestamp.getDate()} ${thaiMonths[timestamp.getMonth()]} ${thaiYear}`;
    const thaiTime = `${timestamp.getHours().toString().padStart(2, '0')}.${timestamp.getMinutes().toString().padStart(2, '0')}`;
    const thaiDateTime = `${thaiDate} ${thaiTime}`;

    // Format the provided date (data.date) to Thai date format
    const inputDate = new Date(data.date);
    const inputThaiDate = `${inputDate.getDate()} ${thaiMonths[inputDate.getMonth()]} ${inputDate.getFullYear() + 543}`;

    // บันทึกข้อมูลใน book_id
    bookSheet.appendRow([timestamp, data.date, data.from, data.to, data.subject, data.action]);

    // Handle book_sent sheet - รับเลขถัดไป
    const lastRow = sentSheet.getLastRow();
    let lastNumber = 0;
    if (lastRow > 0) {
      const lastValue = sentSheet.getRange(lastRow, 1).getValue();
      lastNumber = parseInt(lastValue, 10) || 0;
    }

    const newNumber = (lastNumber + 1).toString().padStart(4, '0');
    sentSheet.appendRow([newNumber, thaiDateTime, inputThaiDate, data.from, data.to, data.subject, data.action]);

    return { success: true, number: newNumber };

  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึกข้อมูล กรุณาลองใหม่อีกครั้ง' };
  } finally {
    // ปล่อย lock เสมอ
    try {
      lockSheet.getRange(1, 1, 1, 2).clearContent();
    } catch (e) {
      // ถ้าไม่สามารถปล่อย lock ได้ ก็ไม่เป็นไร เพราะมี timeout อยู่แล้ว
    }
  }
}

function getSentOptions() {
  const ss = SpreadsheetApp.openById('1LHV8jSSsQ7TAc1Emi_888jliem003l2ys4EIEqMh_GQ');
  const sheet = ss.getSheetByName('data');

  if (!sheet) throw new Error('Sheet data not found');

  const values = sheet.getRange('A2:A').getValues().flat().filter(String);
  return values;
}

function getFromOptions() {
  const ss = SpreadsheetApp.openById('1LHV8jSSsQ7TAc1Emi_888jliem003l2ys4EIEqMh_GQ');
  const sheet = ss.getSheetByName('data');

  if (!sheet) throw new Error('Sheet data not found');

  const values = sheet.getRange('B2:B').getValues().flat().filter(String);
  return values;
}