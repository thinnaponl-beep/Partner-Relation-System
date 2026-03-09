// =========================================================================
// 🌟 ฟังก์ชันแจ้งเตือน KPI (3-Tier Alert) ส่งเข้า Slack 🌟
// =========================================================================
const SLACK_WEBHOOK_URL = 'https://hooks.slack.com/services/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'
const WEB_APP_URL = 'https://script.google.com/a/beneat.co/macros/s/AKfycbxswzmocLIBJ5JrCMuhOZQuiw4vo_PDq3weBU5OKcziJyuEJ21dTtqQOqisoi6Ab195/exec?page=hold'; // ⚠️ เปลี่ยนเป็น URL Web App ของคุณ (ลิงก์ที่ใช้เข้าหน้า Hold)

// ใส่ชื่อชีตฐานข้อมูล Hold ให้ตรงกับของจริง (เช่น 'Database_Hold' หรือ 'Hold')
const HOLD_SHEET_ACTUAL_NAME = 'Database_Hold'; // ⚠️ เปลี่ยนให้ตรงกับชื่อชีตของคุณถ้าไม่ใช่ชื่อนี้

function sendDailyKpiAlertToSlack() {
  if (!SLACK_WEBHOOK_URL || SLACK_WEBHOOK_URL === 'YOUR_SLACK_WEBHOOK_URL_HERE') {
    Logger.log('❌ Error: ยังไม่ได้ตั้งค่า SLACK_WEBHOOK_URL');
    return;
  }

  const holdData = getHoldDataForSlack(); 
  if (!holdData || holdData.length === 0) {
    Logger.log('⚠️ ไม่พบข้อมูลในระบบ หรือดึงข้อมูลไม่สำเร็จ');
    return;
  }

  Logger.log(`✅ ดึงข้อมูลสำเร็จ: พบแม่บ้านทั้งหมด ${holdData.length} รายการ (กำลังตรวจสอบเงื่อนไข...)`);

  let criticalList = []; // แดง (<= 5 วัน)
  let urgentList = [];   // ส้ม (6-10 วัน)
  let warningList = [];  // เหลือง (11-20 วัน)

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  holdData.forEach(row => {
    // แจ้งเตือนเฉพาะคนที่สถานะ "เปิดระบบ" เท่านั้น
    if (row.status !== 'เปิดระบบ') return;
    if (!row.reactivationDate) return;

    let openDate = parseDateForSort(row.reactivationDate);
    if (!openDate) return;
    
    // แปลงกลับเป็น Object Date เพื่อบวกวัน
    let d = new Date(openDate);
    d.setDate(d.getDate() + 30); // ครบกำหนด 30 วัน
    d.setHours(0, 0, 0, 0);

    let daysLeft = Math.ceil((d.getTime() - today.getTime()) / (1000 * 3600 * 24));

    // นับจำนวนงานที่รับไปแล้ว
    let jobsCount = 0;
    if (row.kpiJobs && Array.isArray(row.kpiJobs)) {
      jobsCount = row.kpiJobs.filter(j => j.booking && String(j.booking).trim() !== "").length;
    }

    let daysText = daysLeft < 0 ? `*เกินกำหนดมา ${Math.abs(daysLeft)} วัน*` : `เหลือ *${daysLeft} วัน*`;
    let detailMsg = `• รหัส ${row.maidCode || '-'} - ${row.name} (รับไปแล้ว ${jobsCount}/4 งาน) - ${daysText}`;

    // จัดกลุ่มตามระยะเวลา 3 ระดับ
    if (daysLeft <= 7) {
      criticalList.push(detailMsg);
      Logger.log(`- พบกลุ่มวิกฤต: ${row.name} (${daysLeft} วัน)`);
    } else if (daysLeft > 8 && daysLeft <= 14) {
      urgentList.push(detailMsg);
      Logger.log(`- พบกลุ่มเร่งด่วน: ${row.name} (${daysLeft} วัน)`);
    } else if (daysLeft > 15 && daysLeft <= 21) {
      warningList.push(detailMsg);
      Logger.log(`- พบกลุ่มแจ้งเตือน: ${row.name} (${daysLeft} วัน)`);
    }
  });

  // ถ้าไม่มีใครเข้าข่ายเลย ไม่ต้องส่งข้อความไปกวน
  if (criticalList.length === 0 && urgentList.length === 0 && warningList.length === 0) {
    Logger.log('👌 All clear! วันนี้ไม่มีเคสที่ต้องแจ้งเตือนเข้า Slack (ทุกคนเหลือเวลา > 20 วัน หรือยังไม่เปิดระบบ)');
    return;
  }

  // ประกอบร่างข้อความ (Markdown สไตล์ Slack)
  let finalMessage = "🔔 แจ้งเตือนติดตามรับงาน (คุณแม่บ้านกลุ่ม H Return)* 🔔\n\n";

  if (criticalList.length > 0) {
    finalMessage += "🔴 คุณแม่บ้านที่ยังเข้าให้บริการไม่ครบ 2 งาน (เหลือไม่เกิน 7 วันจะถูกปิดระบบ):*\n";
    finalMessage += criticalList.join("\n") + "\n\n";
  }

  if (urgentList.length > 0) {
    finalMessage += "🟠 คุณแม่บ้านที่ยังเข้าให้บริการไม่ครบ 2 งาน (เหลือไม่เกิน 14 วันจะถูกปิดระบบ):*\n";
    finalMessage += urgentList.join("\n") + "\n\n";
  }

  if (warningList.length > 0) {
    finalMessage += "🟡 คุณแม่บ้านที่ยังเข้าให้บริการไม่ครบ 2 งาน (เหลือไม่เกิน 21 วันจะถูกปิดระบบ):*\n";
    finalMessage += warningList.join("\n") + "\n\n";
  }

  // ใช้รูปแบบแนบลิงก์ของ Slack <URL|ข้อความที่จะแสดง>
  finalMessage += `_*👉 กรุณาตรวจสอบและติดตาม :*_ <${WEB_APP_URL}|ติดตามคุณแม่บ้านกลุ่ม H >`;

  // ฟังก์ชันยิง Webhook ไป Slack
  postToSlack(finalMessage);
}

// ฟังก์ชันสำหรับยิง API ไปยัง Slack
function postToSlack(message) {
  const payload = {
    "text": message
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
    Logger.log("✅ ส่งข้อความแจ้งเตือนเข้า Slack สำเร็จ!");
  } catch (e) {
    Logger.log("❌ Error sending to Slack: " + e.toString());
  }
}

// ฟังก์ชัน Helper เพื่อดึงข้อมูล แบบ Safe Mode
function getHoldDataForSlack() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // พยายามใช้ชื่อชีตจากค่าคงที่ก่อน ถ้าไม่มีให้ใช้ชื่อที่ระบุเอง
  let sheetName = typeof HOLD_SHEET_NAME !== 'undefined' ? HOLD_SHEET_NAME : HOLD_SHEET_ACTUAL_NAME;
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log(`❌ Error: ไม่พบชีตที่ชื่อ '${sheetName}' กรุณาตรวจสอบชื่อชีต`);
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // ดึงข้อมูล 19 คอลัมน์ (A ถึง S) เผื่อไว้
  const values = sheet.getRange(2, 1, lastRow - 1, 19).getDisplayValues();
  let data = values.map(row => {
      let kpiJobs = [];
      try { if (row[16]) kpiJobs = JSON.parse(row[16]); } catch(e) {}

      return {
          maidCode: row[1], // คอลัมน์ B
          name: row[2],     // คอลัมน์ C
          status: row[14],  // คอลัมน์ O (ตรวจสอบว่าชีตจริง สถานะอยู่คอลัมน์ O ใช่หรือไม่)
          reactivationDate: row[15], // คอลัมน์ P (วันที่เปิดระบบใหม่)
          kpiJobs: kpiJobs
      };
  });
  return data;
}

// ฟังก์ชันแปลงวันที่
function parseDateForSort(dateStr) {
  if (!dateStr) return 0;
  if (dateStr instanceof Date) return dateStr.getTime();
  
  if (typeof dateStr === 'string') {
      if (dateStr.match(/^\d{4}-\d{2}-\d{2}/)) {
          const parts = dateStr.substring(0, 10).split('-');
          let y = parseInt(parts[0], 10);
          if (y > 2400) y -= 543;
          return new Date(y, parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)).getTime();
      }
      if (dateStr.match(/^\d{1,2}\/\d{1,2}\/\d{4}/)) {
          const parts = dateStr.split(/[/\s:]/);
          let d = parseInt(parts[0], 10);
          let m = parseInt(parts[1], 10) - 1;
          let y = parseInt(parts[2], 10);
          if (y > 2400) y -= 543;
          return new Date(y, m, d).getTime();
      }
  }
  return 0;
}
