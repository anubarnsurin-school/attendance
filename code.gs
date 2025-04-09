/**
 * ระบบบันทึกเวลาเรียนด้วย Google Apps Script
 * เวอร์ชันสำหรับ GitHub
 */

// ใช้ Config จากไฟล์ config.gs
var CONFIG = getConfig();

function checkFileStructure() {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheets = ss.getSheets();

  if (sheets.length < 4) {
    throw new Error("ไฟล์นี้มีโครงสร้างไม่ถูกต้อง\nควรมีอย่างน้อย 4 แผ่นงาน:\n1. password\n2. setting\n3. Total\n4. StdDonot");
  }

  const requiredSheets = ["password", "setting", "Total", "StdDonot"];
  requiredSheets.forEach((sheetName, index) => {
    if (sheets[index].getName() !== sheetName) {
      throw new Error(`แผ่นงานที่ ${index + 1} ควรชื่อ '${sheetName}'`);
    }
  });

  return "โครงสร้างไฟล์ถูกต้อง";
}

function doGet() {
  try {
    checkFileStructure();
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('บันทึกเวลาเรียน')
      .setFaviconUrl(CONFIG.faviconUrl)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    return HtmlService.createHtmlOutput(`
      <h1>เกิดข้อผิดพลาด</h1>
      <p>${e.message}</p>
      <p>กรุณาตรวจสอบโครงสร้างไฟล์และลองอีกครั้ง</p>
    `);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getClassList() {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const classList = [];

  // เริ่มจากแผ่นที่ sheetNum เป็นต้นไป
  for (let i = CONFIG.sheetNum; i < ss.getNumSheets(); i++) {
    const sheet = ss.getSheets()[i];
    classList.push(sheet.getName());
  }

  return classList;
}

function getStudentData(className) {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName(className);
  
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return data.flat();
}

// ฟังก์ชันอื่นๆ ที่เหลือให้คงเดิม แต่เปลี่ยนจากการใช้ SpreadsheetApp.getActiveSpreadsheet()////////////////////
// เป็น SpreadsheetApp.openById(CONFIG.sheetId) ทุกที่

// ปรับปรุงฟังก์ชัน saveAttendance เพื่อรองรับการอัปเดต
function saveAttendance(data, date, className, isUpdate = false) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.sheetId).getSheetByName(className);///////////////////////
    if (!sheet) throw new Error("ไม่พบชีตสำหรับชั้นเรียนนี้");

    // บันทึกข้อมูลลงชีต
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let dateIndex = headers.indexOf(date);

    if (dateIndex === -1) {
      dateIndex = headers.length;
      sheet.getRange(1, dateIndex + 1).setValue(date);
    }

    for (let i = 0; i < data.length; i++) {
      sheet.getRange(i + 2, dateIndex + 1).setValue(data[i]);
    }

    deleteColumn(sheet);

    // ส่งรายงานผ่าน Telegram
    try {
      const reportResult = SlidePDFJPG(className);
      Logger.log(reportResult);
      return {
        success: true,
        message: isUpdate ? "อัปเดตข้อมูลและส่งรายงานเรียบร้อยแล้ว" : "บันทึกข้อมูลและส่งรายงานเรียบร้อยแล้ว",
        report: reportResult
      };
    } catch (e) {
      Logger.log("เกิดข้อผิดพลาดขณะส่งรายงาน: " + e.message);
      return {
        success: true,
        message: isUpdate ? "อัปเดตข้อมูลเรียบร้อย แต่ส่งรายงานไม่สำเร็จ" : "บันทึกข้อมูลเรียบร้อย แต่ส่งรายงานไม่สำเร็จ",
        error: e.message
      };
    }

  } catch (e) {
    Logger.log("Error in saveAttendance: " + e.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + e.message
    };
  }
}

function deleteColumn(sheet) {
  var lastColumn = sheet.getLastColumn();
  var nameList = [];

  for (var col = lastColumn; col >= 2; col--) {
    var name = sheet.getRange(1, col).getDisplayValue();

    if (nameList.includes(name)) {
      sheet.deleteColumn(col);
    } else {
      nameList.push(name);
    }
  }
}


function verifyPassword(className, inputPassword) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.sheetId);//////////////////////
    const passwordSheet = ss.getSheetByName("password");

    if (!passwordSheet) {
      return JSON.stringify({
        success: false,
        message: "ไม่พบแผ่นงานรหัสผ่านในระบบ"
      });
    }

    const data = passwordSheet.getDataRange().getValues();
    const headers = data[0];
    const classNameCol = headers.indexOf("ชั้นเรียน");
    const passwordCol = headers.indexOf("รหัสผ่าน");

    // ตรวจสอบโครงสร้างชีต
    if (classNameCol === -1 || passwordCol === -1) {
      return JSON.stringify({
        success: false,
        message: "รูปแบบแผ่นงานรหัสผ่านไม่ถูกต้อง\nกรุณาตรวจสอบหัวคอลัมน์"
      });
    }

    // ค้นหารหัสผ่านสำหรับชั้นเรียนนี้
    for (let i = 1; i < data.length; i++) {
      if (data[i][classNameCol] && data[i][classNameCol].trim() === className.trim()) {
        if (data[i][passwordCol] == inputPassword) {
          return JSON.stringify({
            success: true,
            message: "รหัสผ่านถูกต้อง"
          });
        }
        return JSON.stringify({
          success: false,
          message: "รหัสผ่านไม่ถูกต้อง"
        });
      }
    }

    return JSON.stringify({
      success: false,
      message: `ไม่พบรหัสผ่านสำหรับชั้น ${className}`
    });

  } catch (e) {
    return JSON.stringify({
      success: false,
      message: `เกิดข้อผิดพลาด: ${e.message}`
    });
  }
}

/**
 * ขอบคุณคุณครูสมพงษ์  โพคาศรี (ไฟล์ต้นฉบับจากไลน์แจ้งเตือน)
 * ฟังก์ชันหลักสำหรับสร้างรายงานและส่งผ่าน Telegram
 * @param {string} className - ชื่อชั้นเรียน (ต้องตรงกับชื่อชีต)
 */
function SlidePDFJPG(className) {
  try {
    // 1. ตรวจสอบและเตรียมข้อมูลพื้นฐาน
    // var className = "ม.4/1"; // เปลี่ยนเป็นชื่อชีตที่ต้องการทดสอบ

    var sheet = SpreadsheetApp.openById(CONFIG.sheetId).getSheetByName(className);//////////////////
    if (!sheet) throw new Error("ไม่พบชีต '" + className + "'");

    // 2. ดึงข้อมูลจากชีต
    var data = sheet.getDataRange().getValues();

    // 3. ดึงค่า Token และ Chat ID จากคอลัมน์ Q แถวที่ 8-9
    var telegramBotToken = data[1][17]; // แถว 2 คอลัมน์ Q
    var telegramChatId = data[2][17];   // แถว 3 คอลัมน์ Q


    // 4. ตรวจสอบค่าที่จำเป็น
    if (!telegramBotToken || !telegramChatId) {
      throw new Error("กรุณากรอก Bot Token (r8) และ Chat ID (r9) ในชีต " + className);
    }

    if (!telegramBotToken.match(/^\d+:[a-zA-Z0-9_-]+$/)) {
      throw new Error("รูปแบบ Bot Token ไม่ถูกต้อง");
    }

    // แก้ไขบรรทัด 216-218 ในฟังก์ชัน SlidePDFJPG
    var telegramChatId = String(data[2][17] || "").trim(); // ดึงค่าและแปลงเป็น String


    if (!telegramChatId) {
      throw new Error("ไม่พบ Chat ID ในเซลล์ r9");
    }


    if (!telegramChatId.match(/^-?\d+$/)) {  // ใช้ Regular Expression ที่ถูกต้อง
      throw new Error("รูปแบบ Chat ID ไม่ถูกต้อง\nต้องเป็นตัวเลขเท่านั้น (เช่น 123456 หรือ -100123456)");
    }


    Logger.log("กำลังเตรียมส่งรายงานสำหรับชั้น " + className + " ไปยัง Telegram...");


    // 5. เตรียมข้อมูลวันที่
    var strWeek = ["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"];
    var d = new Date();
    var outputDate = 'วัน' + strWeek[d.getDay()] + ' ที่ ' +
      Utilities.formatDate(d, 'GMT+7', 'dd/MM/') +
      (d.getFullYear() + 543);
    var dateReport = Utilities.formatDate(d, 'GMT+7', 'dd/MM/') +
      (d.getFullYear() + 543);


    // 6. เตรียมไฟล์ชั่วคราว
    var tempFolder = DriveApp.getFolderById(folderId);
    var templateId = slideId;
    var templateFile = DriveApp.getFileById(templateId);
    var copyFile = templateFile.makeCopy('รายงาน_' + className + '_' + new Date().getTime(), tempFolder);
    var copyId = copyFile.getId();
    var copyDoc = SlidesApp.openById(copyId);


    // 7. ประมวลผลข้อมูลนักเรียน
    var studentData = data.slice(1); // ลบแถวหัวข้อออก
    var totalStudents = studentData.length;
    var presentCount = 0;
    var absentCount = 0;
    var tellCount = 0;
    var leaveCount = 0;


    for (var i = 0; i < studentData.length; i++) {
      var student = studentData[i];

      // อัพเดตข้อมูลในสไลด์
      copyDoc.getSlides()[0].getTables()[1].getCell(i, 0).getText().setText(student[0]); // ชื่อ-สกุล
      copyDoc.getSlides()[0].getTables()[2].getCell(i, 0).getText().setText(student[19]); // สถานนะ
      copyDoc.getSlides()[0].getTables()[3].getCell(i, 0).getText().setText(student[3]); // สถิติขาด
      copyDoc.getSlides()[0].getTables()[4].getCell(i, 0).getText().setText(student[4]); // สถิติลา
      copyDoc.getSlides()[0].getTables()[5].getCell(i, 0).getText().setText(student[5]); // สถิติเต็ม
      copyDoc.getSlides()[0].getTables()[6].getCell(i, 0).getText().setText(student[6]); // ไม่ร่วมกิจกรรม
      copyDoc.getSlides()[0].getTables()[7].getCell(i, 0).getText().setText(student[7]); // ไม่ร่วมกิจกรรม

      // นับจำนวนสถานะ
      if (student[19] === 'มาเรียน') {
        presentCount++;
      } else if (student[19] === 'ขาด') {
        absentCount++;
      } else if (student[19] === 'ลา') {
        tellCount++;
      } else {
        leaveCount++;
      }
    }


    // 8. คำนวณสถิติ
    var presentPercent = ((presentCount + leaveCount) * 100 / totalStudents).toFixed(2) + '%';
    var absentPercent = (absentCount * 100 / totalStudents).toFixed(2) + '%';
    var tellPercent = (tellCount * 100 / totalStudents).toFixed(2) + '%';
    var leavePercent = (leaveCount * 100 / totalStudents).toFixed(2) + '%';


    // 9. อัพเดตข้อมูลในสไลด์
    copyDoc.replaceAllText("<<date>>", outputDate);
    copyDoc.replaceAllText("<<ชั้น>>", "ชั้น " + className);
    copyDoc.replaceAllText("<<ars01>>", totalStudents + "(100%)");
    copyDoc.replaceAllText("<<ars02>>", (presentCount + leaveCount) + "(" + presentPercent + ")");
    copyDoc.replaceAllText("<<ars03>>", absentCount + "(" + absentPercent + ")");
    copyDoc.replaceAllText("<<ars04>>", tellCount + "(" + tellPercent + ")");
    copyDoc.replaceAllText("<<ars05>>", leaveCount + "(" + leavePercent + ")");
    copyDoc.replaceAllText("<<ars06>>", (totalStudents - presentCount) + "(" + (100 - parseFloat(presentPercent)).toFixed(2) + "%)");


    // 10. บันทึกและปิดไฟล์
    copyDoc.saveAndClose();


    // 11. สร้างรูปภาพจากสไลด์
    var slide = copyDoc.getSlides()[0];
    var thumbnailUrl = Slides.Presentations.Pages.getThumbnail(
      copyId,
      slide.getObjectId(),
      { "thumbnailProperties.mimeType": "PNG" }
    ).contentUrl;

    var imageBlob = UrlFetchApp.fetch(thumbnailUrl).getAs(MimeType.JPEG);
    var imageFile = tempFolder.createFile(imageBlob.setName('รายงาน_' + className + '-' + dateReport + '.jpg'));
    var imageUrl = 'https://lh5.googleusercontent.com/d/' + imageFile.getId(); // URL ของรูปภาพ
    console.log(imageUrl)

    // 12. เตรียมข้อความแจ้งเตือน
    var message = `📊 รายงานการมาเรียน : ชั้น ${className}
📅 ${outputDate}
👨‍🎓 นักเรียนทั้งหมด : ${totalStudents} คน
✅ มาเรียน : ${(presentCount + leaveCount)} คน (${presentPercent})
❌ ขาด : ${absentCount} คน (${absentPercent})
📝 ลา : ${tellCount} คน (${tellPercent})
🤫 หนี : ${leaveCount} คน (${leavePercent})`
+ "\n\n📸 ดาวน์โหลดรูปภาพเต็มขนาด:\n" + imageUrl;




    // 13. ส่งแจ้งเตือนผ่าน Telegram
    var telegramResult = sendTelegramNotify(telegramChatId, telegramBotToken, message, imageFile.getBlob());

    Logger.log("ส่งรายงานเสร็จสิ้น: " + JSON.stringify(telegramResult));


    // 14. ลบไฟล์ชั่วคราว
    copyFile.setTrashed(true);
    // imageFile.setTrashed(true);
    imageFile.setTrashed(false);


    return "ส่งรายงานเรียบร้อยแล้วสำหรับชั้น " + className;


  } catch (e) {
    Logger.log("เกิดข้อผิดพลาด: " + e.message);
    throw e;
  }
}




function createFormData(data, boundary) {
  var payload = [];

  for (var key in data) {
    if (data.hasOwnProperty(key)) {
      if (typeof data[key] === 'object' && data[key].fileName) {
        // สำหรับไฟล์
        payload.push(
          '--' + boundary,
          'Content-Disposition: form-data; name="' + key + '"; filename="' + data[key].fileName + '"',
          'Content-Type: ' + data[key].mimeType,
          '',
          Utilities.newBlob(data[key].content, data[key].mimeType).getDataAsString()
        );
      } else {
        // สำหรับข้อมูลปกติ
        payload.push(
          '--' + boundary,
          'Content-Disposition: form-data; name="' + key + '"',
          '',
          String(data[key])
        );
      }
    }
  }

  payload.push('--' + boundary + '--');
  return payload.join('\r\n');
}



/**
 * ฟังก์ชันส่งการแจ้งเตือนผ่าน Telegram
 * @param {string} chatId - ID ของแชทหรือกลุ่ม
 * @param {string} botToken - Token ของบอท Telegram
 * @param {string} message - ข้อความที่จะส่ง
 * @param {Blob} imageBlob - ไฟล์รูปภาพ (optional)
 */
function sendTelegramNotify(chatId, botToken, message, imageBlob) {
  try {
    var telegramUrl = "https://api.telegram.org/bot" + botToken;
    var options = {
      "muteHttpExceptions": true
    };


    // ตรวจสอบและแปลงค่า
    chatId = String(chatId).trim();
    botToken = String(botToken).trim();


    if (imageBlob) {
      // ส่งรูปภาพ (ใช้ FormData)
      var formData = {
        "chat_id": chatId,
        "caption": message,
        "parse_mode": "HTML",
        "photo": imageBlob
      };

      options.method = "post";
      options.payload = formData;

      var response = UrlFetchApp.fetch(telegramUrl + "/sendPhoto", options);
    } else {
      // ส่งข้อความธรรมดา
      options.method = "post";
      options.payload = JSON.stringify({
        "chat_id": chatId,
        "text": message,
        "parse_mode": "HTML"
      });
      options.contentType = "application/json";

      var response = UrlFetchApp.fetch(telegramUrl + "/sendMessage", options);
    }


    // ตรวจสอบผลลัพธ์
    var responseText = response.getContentText();
    if (!responseText) {
      throw new Error("ไม่ได้รับข้อมูลตอบกลับ");
    }

    var result = JSON.parse(responseText);

    if (!result.ok) {
      throw new Error("Telegram API Error: " + responseText);
    }

    return result;

  } catch (e) {
    Logger.log("รายละเอียดข้อผิดพลาด:\n" + e.stack);
    throw new Error("ส่งข้อความไม่สำเร็จ: " + e.message);
  }
}




/**
 * ฟังก์ชันทดสอบการเชื่อมต่อ Telegram
 */
function testTelegramConnection() {
  var className = "ม.1/1"; // เปลี่ยนเป็นชื่อชีตที่ต้องการทดสอบ
  var sheet = SpreadsheetApp.openById(CONFIG.sheetId).getSheetByName(className);///////////////////////

  if (!sheet) {
    Logger.log("ไม่พบชีต '" + className + "'");
    return;
  }

  var data = sheet.getDataRange().getValues();
  var botToken = data[1][17]; // r8
  var chatId = data[2][17];   // r9

  if (!botToken || !chatId) {
    Logger.log("กรุณากรอก Bot Token (r8) และ Chat ID (r9) ในชีต " + className);
    return;
  }

  var testMessage = "🔔 ทดสอบการเชื่อมต่อ Telegram จาก Google Sheets\n" +
    "ชั้นเรียน: " + className + "\n" +
    "เวลา: " + new Date();

  try {
    var result = sendTelegramNotify(chatId, botToken, testMessage);
    Logger.log("ทดสอบส่งข้อความสำเร็จ: " + JSON.stringify(result));
  } catch (e) {
    Logger.log("การทดสอบล้มเหลว: " + e.message);
  }
}

// เพิ่มฟังก์ชันดึงข้อมูลการบันทึกเดิม
function getExistingAttendance(className, date) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.sheetId).getSheetByName(className.trim());/////////////////////////////
    if (!sheet) throw new Error("ไม่พบชีตสำหรับชั้นเรียนนี้");
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const dateIndex = headers.indexOf(date);
    
    if (dateIndex === -1) {
      return { exists: false };
    }
    
    const studentNames = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const attendanceData = sheet.getRange(2, dateIndex + 1, sheet.getLastRow() - 1, 1).getValues().flat();
    
    // ตรวจสอบว่ามีข้อมูลจริงหรือไม่
    const hasData = attendanceData.some(cell => cell.toString().trim() !== '');
    
    return {
      exists: hasData,
      studentNames: studentNames,
      attendanceData: attendanceData
    };
  } catch (e) {
    Logger.log("Error in getExistingAttendance: " + e.message);
    throw e;
  }
}
