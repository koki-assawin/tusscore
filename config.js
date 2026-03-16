// ==================== TussScore Configuration ====================
// แก้ไขค่าในไฟล์นี้เพื่อตั้งค่าระบบ

const TUSSCORE_CONFIG = {
  // Google Apps Script URL สำหรับเชื่อมต่อฐานข้อมูล
  // แก้ไข URL นี้เมื่อต้องการเปลี่ยน Google Sheets
  scriptUrl: 'https://script.google.com/macros/s/AKfycbzT21vAr_AXxR_pOQKeMsm-YZ8y0bCwJspMls70_oCL92bBXHKAH7K16V2JdCMHYly2/exec',
};

// ตั้งค่า localStorage อัตโนมัติ เพื่อให้ระบบทำงานได้ทันที
if (TUSSCORE_CONFIG.scriptUrl && !localStorage.getItem('tusscore_script_url')) {
  localStorage.setItem('tusscore_script_url', TUSSCORE_CONFIG.scriptUrl);
}
