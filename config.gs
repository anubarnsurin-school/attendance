/**
 * ค่าคอนฟิกที่สามารถกำหนดผ่าน Script Properties
 * วิธีตั้งค่า: File > Project properties > Script properties
 * 
 * ต้องตั้งค่าดังนี้:
 * - SHEET_ID: ID ของ Google Spreadsheet
 * - SLIDE_TEMPLATE_ID: ID ของ Google Slides template
 * - REPORT_FOLDER_ID: ID ของโฟลเดอร์เก็บรายงาน
 * - COMPANY_NAME: ชื่อโรงเรียน/สถาบัน
 * - LOGO_URL: URL ของโลโก้
 * - FAVICON_URL: URL ของ favicon
 */
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  
  return {
    sheetId: props.getProperty('1oOqUFyiKU7ukN2LwQbTOahdqiTwnyGI_rzVtQyA378w'),
    slideTemplateId: props.getProperty('1oRFQuEYGDdQ1RhRAh_Y9CqJl0dc6xsW2t1njdsKw9Dg'),
    reportFolderId: props.getProperty('1AVZqIctJeQ7U0vXZDFvr8AiwMXD9r3YJ'),
    companyName: props.getProperty('โรงเรียนเทนมีย์มิตรประชา') || 'โรงเรียนตัวอย่าง',
    logoUrl: props.getProperty('https://sp0125.github.io/tklimage/JO63CP.png') || 'https://via.placeholder.com/150',
    faviconUrl: props.getProperty('https://sp0125.github.io/tklimage/JO63CP.png') || 'https://via.placeholder.com/32',
    sheetNum: 4 // จำนวนชีตเริ่มต้นที่ไม่ใช่ข้อมูลนักเรียน (password, setting, Total, StdDonot)
  };
}
