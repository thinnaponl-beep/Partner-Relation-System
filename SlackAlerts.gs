// =========================================================================
// 🌟 ฟังก์ชันแจ้งเตือน KPI (3-Tier Alert) ส่งเข้า Slack (ดึงข้อมูลจาก Supabase) 🌟
// =========================================================================

// ⚠️ ตรวจสอบ Webhook URL ว่ายังใช้งานได้หรือไม่
var SLACK_WEBHOOK_URL = '';
var WEB_APP_URL = '';

// 🌟 ข้อมูลการเชื่อมต่อ Supabase (ใช้ var เพื่อป้องกัน Error is not defined ข้ามไฟล์)
// อัปเดตฐานข้อมูลเป็นชุดล่าสุดให้ตรงกับฝั่งเว็บไซต์
var SUPABASE_URL = 'https://bnzlooehcmgsqgcwjqcj.supabase.co'; 
var SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJuemxvb2VoY21nc3FnY3dqcWNqIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3NjM5Mzc2MSwiZXhwIjoyMDkxOTY5NzYxfQ.ZswcbFK8tlJBV8ONQKuM8qNpEO4AJ0jVyd0JSfL-fBI';

function sendDailyKpiAlertToSlack() {
  if (!SLACK_WEBHOOK_URL || SLACK_WEBHOOK_URL === 'YOUR_SLACK_WEBHOOK_URL_HERE') {
    Logger.log('❌ Error: ยังไม่ได้ตั้งค่า SLACK_WEBHOOK_URL');
    return;
  }

  const holdData = getHoldDataFromSupabase(); 
  if (!holdData || holdData.length === 0) {
    Logger.log('⚠️ ไม่พบข้อมูลในระบบ Supabase หรือดึงข้อมูลไม่สำเร็จ');
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

    // นับจำนวนงานที่รับไปแล้ว (คอลัมน์ JSONB จาก Supabase จะถูกแปลงเป็น Array อัตโนมัติ)
    let jobsCount = 0;
    if (row.kpiJobs && Array.isArray(row.kpiJobs)) {
      jobsCount = row.kpiJobs.filter(j => j.booking && String(j.booking).trim() !== "").length;
    }

    let daysText = daysLeft < 0 ? `*เกินกำหนดมา ${Math.abs(daysLeft)} วัน*` : `เหลือ *${daysLeft} วัน*`;
    let detailMsg = `• รหัส ${row.maidCode || '-'} - ${row.name} (รับไปแล้ว ${jobsCount}/4 งาน) - ${daysText}`;

    // จัดกลุ่มตามระยะเวลา 3 ระดับ (ให้ตรงกับตัวกรองบนหน้าเว็บ 5, 10, 20 วัน)
    if (daysLeft <= 5) {
      criticalList.push(detailMsg);
      Logger.log(`- พบกลุ่มวิกฤต: ${row.name} (${daysLeft} วัน)`);
    } else if (daysLeft > 5 && daysLeft <= 10) {
      urgentList.push(detailMsg);
      Logger.log(`- พบกลุ่มเร่งด่วน: ${row.name} (${daysLeft} วัน)`);
    } else if (daysLeft > 10 && daysLeft <= 20) {
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
  let finalMessage = "🔔 *แจ้งเตือนติดตามรับงาน (คุณแม่บ้านกลุ่ม H Return)* 🔔\n\n";

  if (criticalList.length > 0) {
    finalMessage += "🔴 *กลุ่มวิกฤต (เหลือเวลาไม่เกิน 5 วัน หรือเกินกำหนด):*\n";
    finalMessage += criticalList.join("\n") + "\n\n";
  }

  if (urgentList.length > 0) {
    finalMessage += "🟠 *กลุ่มเร่งด่วน (เหลือเวลา 6-10 วัน):*\n";
    finalMessage += urgentList.join("\n") + "\n\n";
  }

  if (warningList.length > 0) {
    finalMessage += "🟡 *กลุ่มแจ้งเตือน (เหลือเวลา 11-20 วัน):*\n";
    finalMessage += warningList.join("\n") + "\n\n";
  }

  // ใช้รูปแบบแนบลิงก์ของ Slack <URL|ข้อความที่จะแสดง>
  finalMessage += `_*👉 กรุณาตรวจสอบและติดตาม :*_ <${WEB_APP_URL}|คลิกที่นี่เพื่อเปิดหน้าระบบ PRS >`;

  // ฟังก์ชันยิง Webhook ไป Slack
  postToSlack(finalMessage);
}

// ฟังก์ชันยิง REST API เพื่อดึงข้อมูลจาก Supabase 
function getHoldDataFromSupabase() {
  const url = `${SUPABASE_URL}/rest/v1/prs_hold?select=maid_code,full_name,status,reactivation_date,kpi_jobs`;
  
  const options = {
    method: 'get',
    headers: {
      'apikey': SUPABASE_KEY,
      'Authorization': `Bearer ${SUPABASE_KEY}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200 || response.getResponseCode() === 201) {
      const data = JSON.parse(response.getContentText());
      return data.map(row => ({
        maidCode: row.maid_code,
        name: row.full_name,
        status: row.status,
        reactivationDate: row.reactivation_date,
        kpiJobs: row.kpi_jobs || [] 
      }));
    } else {
      Logger.log(`❌ Supabase Error: ${response.getContentText()}`);
    }
  } catch (e) {
    Logger.log(`❌ Fetch Exception: ${e.toString()}`);
  }
  return [];
}

// ฟังก์ชันสำหรับยิง API ไปยัง Slack
function postToSlack(message) {
  const payload = {
    "text": message
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true 
  };

  try {
    const response = UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200 || response.getContentText() === "ok") {
      Logger.log("✅ ส่งข้อความแจ้งเตือนเข้า Slack สำเร็จ!");
    } else {
      Logger.log(`❌ ไม่สามารถส่งเข้า Slack ได้ (Code ${responseCode}) - สาเหตุจาก Slack: ${response.getContentText()}`);
      Logger.log(`⚠️ กรุณาตรวจสอบลิงก์ Webhook ของคุณว่าถูกต้องหรือหมดอายุหรือไม่`);
    }
  } catch (e) {
    Logger.log("❌ Exception sending to Slack: " + e.toString());
  }
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
