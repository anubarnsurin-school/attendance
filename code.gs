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

// ฟังก์ชันอื่นๆ ที่เหลือให้คงเดิม แต่เปลี่ยนจากการใช้ SpreadsheetApp.getActiveSpreadsheet()
// เป็น SpreadsheetApp.openById(CONFIG.sheetId) ทุกที่
