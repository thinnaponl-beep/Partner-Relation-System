const SPREADSHEET_ID = ""; // *** กรุณาใส่ ID ของ Google Sheet ที่นี่ ***
const SHEET_NAME = "Database_Master";
const ANNUAL_SHEET_NAME = "Database_Annual";
const ONBOARD_SHEET_NAME = "Database_Onboard"; 
const FIRSTBK_SHEET_NAME = "Database_Firstbk"; 
const CONFIG_SHEET_NAME = "Config";

// --- Routing ---
function doGet(e) {
  let page = e.parameter.page || 'home'; // *** เปลี่ยน Default เป็น 'home' ***
  let html;
  const user = Session.getActiveUser().getEmail();

  // กำหนดเส้นทาง
  if (page === 'home') {
    html = HtmlService.createTemplateFromFile('Home'); // หน้า Landing Page ใหม่
  } else if (page === 'onboard') { 
    html = HtmlService.createTemplateFromFile('Onboard'); // ติดตามคุณแม่บ้านเปิดระบบ
  } else if (page === 'firstbk') { 
    html = HtmlService.createTemplateFromFile('Firstbk'); // ติดตามรับงานแรก
  } else if (page === 'index') {
    html = HtmlService.createTemplateFromFile('index'); // ตรวจประวัติคุณแม่บ้านใหม่ (หน้า Master เดิม)
  } else if (page === 'year_verify') {
    html = HtmlService.createTemplateFromFile('YearVerify'); // ตรวจประวัติรายปี
  } else if (page === 'config') {
    if (!isUserAdmin(user)) return HtmlService.createHtmlOutput("<h3>Access Denied / คุณไม่มีสิทธิ์เข้าถึงหน้านี้</h3><p>กรุณาติดต่อผู้ดูแลระบบ</p>");
    html = HtmlService.createTemplateFromFile('Config'); // ตั้งค่า
  } else {
    html = HtmlService.createTemplateFromFile('Home');
  }

  return html.evaluate()
    .setTitle('Partner Relation Support')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper: เชื่อมต่อ Sheet (Auto-expand applied)
function getSheet(name) {
  let ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  
  if (name === CONFIG_SHEET_NAME) {
      if (!sheet) {
        sheet = ss.insertSheet(CONFIG_SHEET_NAME);
        sheet.appendRow(["Admin Emails", "Centers", "Results", "FT Statuses", "Onboard Groups", "Onboard Types", "Master Types", "Onboard Centers"]); 
      } else {
        if (sheet.getLastColumn() < 8) {
           sheet.getRange(1, 5).setValue("Onboard Groups");
           sheet.getRange(1, 6).setValue("Onboard Types");
           sheet.getRange(1, 7).setValue("Master Types");
           sheet.getRange(1, 8).setValue("Onboard Centers"); 
        }
      }
  }
  
  if (name === ONBOARD_SHEET_NAME) {
    if (!sheet) {
      sheet = ss.insertSheet(ONBOARD_SHEET_NAME);
      sheet.appendRow([
        "ID", "Training Date", "Maid Code", "Name", "Group", 
        "Phone", "ID Card", "Type", "Latest Followup", "Date 2 (Unused)", "Date 3 (Unused)", 
        "Open Date", "Call Status", "First Job", "Job ID", 
        "Trainer", "History Data (JSON)", "FastTrack Status", "Center", "Skip FastTrack", "Master Type"
      ]);
    } else {
        const currentCols = sheet.getMaxColumns();
        if (currentCols < 21) {
           sheet.insertColumnsAfter(currentCols, 21 - currentCols);
           if(currentCols < 19) sheet.getRange(1, 19).setValue("Center");
           if(currentCols < 20) sheet.getRange(1, 20).setValue("Skip FastTrack");
           if(currentCols < 21) sheet.getRange(1, 21).setValue("Master Type");
        }
    }
  }

  // *** Database_Firstbk Structure ***
  if (name === FIRSTBK_SHEET_NAME) {
    if (!sheet) {
      sheet = ss.insertSheet(FIRSTBK_SHEET_NAME);
      // Added History(22) + Extended checklist cols
      sheet.appendRow([
        "Onboard ID", "Maid Code", "Name", "Phone", "Center", 
        "Booking Code", "Job ID", "Clean Date", "Accept Date", "Status",
        "Check_1_1", "Check_1_2", "Check_1_3", "Check_1_4", "Check_1_5", 
        "Advice", "Officer", "Timestamp", "ReviewScore", "CustomerComment", "ProblemID", "History",
        "Check_1_6", "Check_1_7", 
        "Check_2_1", "Check_2_2", "Check_2_3", "Check_2_4", "Check_2_5", "Check_2_6", "Check_2_7", "Check_2_8"
      ]);
    } else {
        const currentCols = sheet.getMaxColumns();
        if (currentCols < 35) {
            sheet.insertColumnsAfter(currentCols, 35 - currentCols);
        }
    }
  }

  if (name === ANNUAL_SHEET_NAME) {
    if (!sheet) {
      sheet = ss.insertSheet(ANNUAL_SHEET_NAME);
      sheet.appendRow([
        "ID", "Ref Code", "Name", "Group", "ID Card", 
        "Birth Date", "Phone", "Consent Status", "Amount", "Outstanding", "Deduction Status",
        "Channel", "Status Process", "Result Date", "Result", "Last Followup",
        "Officer Email", "Submit Date", "Note", "Export Status"
      ]);
    } else {
        const currentCols = sheet.getMaxColumns();
        if (currentCols < 20) {
           sheet.insertColumnsAfter(currentCols, 20 - currentCols);
        }
    }
  }

  if (name === SHEET_NAME) {
      if(!sheet) {
        sheet = ss.insertSheet(SHEET_NAME);
        sheet.appendRow(["ID", "Ref Code", "Name", "ID Card", "Phone", "Training Date", "Submit Date", "Officer", "Center", "Result Date", "Result", "FT Status", "Note", "Export Status", "Type"]);
      }
  }
  
  return sheet;
}

// --- Config Logic ---
function getConfigs() {
  const sheet = getSheet(CONFIG_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { admins: [], centers: [], results: [], ftStatuses: [], onboardGroups: [], onboardTypes: [], masterTypes: [], onboardCenters: [] };
  
  const maxCols = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, lastRow - 1, maxCols).getValues();
  
  return {
    admins: data.map(r => String(r[0]).trim().toLowerCase()).filter(s => s !== ""),
    centers: data.map(r => String(r[1]).trim()).filter(s => s !== ""),
    results: data.map(r => String(r[2]).trim()).filter(s => s !== ""),
    ftStatuses: data.map(r => String(r[3]).trim()).filter(s => s !== ""),
    onboardGroups: data.map(r => String(r[4]).trim()).filter(s => s !== ""),
    onboardTypes: data.map(r => String(r[5]).trim()).filter(s => s !== ""),
    masterTypes: (maxCols >= 7) ? data.map(r => String(r[6]).trim()).filter(s => s !== "") : [],
    onboardCenters: (maxCols >= 8) ? data.map(r => String(r[7]).trim()).filter(s => s !== "") : [] 
  };
}

function isUserAdmin(email) {
  const configs = getConfigs();
  return configs.admins.includes(String(email).trim().toLowerCase());
}

function getClientConfig() {
  const configs = getConfigs();
  const currentUser = Session.getActiveUser().getEmail();
  return {
    isAdmin: isUserAdmin(currentUser),
    userEmail: currentUser,
    admins: configs.admins,
    centers: configs.centers,
    results: configs.results,
    ftStatuses: configs.ftStatuses,
    onboardGroups: configs.onboardGroups,
    onboardTypes: configs.onboardTypes,
    masterTypes: configs.masterTypes,
    onboardCenters: configs.onboardCenters 
  };
}

function addConfigItem(type, value) {
  if (!value) return { success: false, message: "ค่าว่าง" };
  const sheet = getSheet(CONFIG_SHEET_NAME);
  let colIndex;
  switch(type) {
    case 'admin': colIndex = 1; break;
    case 'center': colIndex = 2; break; 
    case 'result': colIndex = 3; break;
    case 'ftStatus': colIndex = 4; break;
    case 'onboardGroup': colIndex = 5; break;
    case 'onboardType': colIndex = 6; break;
    case 'masterType': colIndex = 7; break;
    case 'onboardCenter': colIndex = 8; break; 
    default: return { success: false, message: "Invalid type" };
  }

  let targetRow = 2;
  while (sheet.getRange(targetRow, colIndex).getValue() !== "") targetRow++;
  sheet.getRange(targetRow, colIndex).setValue(value);
  return { success: true };
}

function removeConfigItem(type, value) {
  const sheet = getSheet(CONFIG_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false };
  
  let colIndex;
  switch(type) {
    case 'admin': colIndex = 1; break;
    case 'center': colIndex = 2; break;
    case 'result': colIndex = 3; break;
    case 'ftStatus': colIndex = 4; break;
    case 'onboardGroup': colIndex = 5; break;
    case 'onboardType': colIndex = 6; break;
    case 'masterType': colIndex = 7; break;
    case 'onboardCenter': colIndex = 8; break;
    default: return { success: false };
  }

  const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
  const values = range.getValues().flat().map(v => String(v).trim().toLowerCase());
  const index = values.indexOf(String(value).trim().toLowerCase());
  if (index !== -1) {
    sheet.getRange(index + 2, colIndex).clearContent();
    const newRange = sheet.getRange(2, colIndex, lastRow - 1, 1);
    const newValues = newRange.getValues().filter(r => r[0] !== "");
    newRange.clearContent();
    if(newValues.length > 0) sheet.getRange(2, colIndex, newValues.length, 1).setValues(newValues);
  }
  return { success: true };
}

function saveConfigOrder(type, newList) {
  const sheet = getSheet(CONFIG_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    let colIndex;
    switch(type) {
      case 'admin': colIndex = 1; break;
      case 'center': colIndex = 2; break;
      case 'result': colIndex = 3; break;
      case 'ftStatus': colIndex = 4; break;
      case 'onboardGroup': colIndex = 5; break;
      case 'onboardType': colIndex = 6; break;
      case 'masterType': colIndex = 7; break;
      case 'onboardCenter': colIndex = 8; break;
      default: return { success: false, message: "Invalid type" };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
       sheet.getRange(2, colIndex, lastRow - 1, 1).clearContent();
    }
    
    if (newList && newList.length > 0) {
       const dataToWrite = newList.map(item => [item]);
       sheet.getRange(2, colIndex, dataToWrite.length, 1).setValues(dataToWrite);
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// --- MASTER DATA LOGIC ---
function getInitialData(filterVal) {
  const sheet = getSheet(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const currentUser = Session.getActiveUser().getEmail();
  let data = [];
  
  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 15).getDisplayValues(); 
    data = values.reduce((acc, row, index) => {
      if (!row[1] && !row[2]) return acc;
      if (filterVal && !isDateInMonth(row[6], filterVal)) return acc;
      acc.push({
        rowIndex: index + 2, id: row[0], code: row[1], name: row[2], idCard: row[3],
        phone: row[4], trainingDate: row[5], submitDate: row[6], officer: row[7],
        center: row[8], resultDate: row[9], result: row[10], ftStatus: row[11],
        note: row[12], exportStatus: row[13], type: row[14]
      });
      return acc;
    }, []);
    data.sort((a, b) => parseDateForSort(a.submitDate) - parseDateForSort(b.submitDate));
  }
  const configs = getClientConfig();
  return { 
    currentUser: currentUser, 
    isAdmin: configs.isAdmin, 
    centers: configs.centers, 
    results: configs.results, 
    ftStatuses: configs.ftStatuses, 
    masterTypes: configs.masterTypes,
    data: data 
  };
}

function saveData(formData) {
  const sheet = getSheet(SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    let rowNumber;
    let newId = formData.id;
    let currentExportStatus = ""; 
    const lastRow = sheet.getLastRow();

    if (formData.id) {
      if (lastRow < 2) throw new Error("Database Empty");
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat().map(id => String(id).trim());
      const index = ids.indexOf(String(formData.id).trim());
      if (index === -1) throw new Error("ID not found");
      rowNumber = index + 2;
      const currentValues = sheet.getRange(rowNumber, 12, 1, 3).getValues()[0]; 
      if (currentValues[0] !== formData.ftStatus) currentExportStatus = ""; else currentExportStatus = currentValues[2];
    } else {
      let maxId = 0;
      if (lastRow >= 2) {
         const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
         existingIds.forEach(id => { let num = Number(id); if (!isNaN(num) && num > maxId) maxId = num; });
      }
      newId = (maxId + 1).toString();
      rowNumber = lastRow + 1;
    }

    const rowData = [
      newId, formData.code, formData.name, "'"+formData.idCard, "'"+formData.phone,
      formatDateForSheet(formData.trainingDate), formatDateForSheet(formData.submitDate), 
      formData.officer, formData.center, formatDateForSheet(formData.resultDate), 
      formData.result, formData.ftStatus, formData.note, currentExportStatus, formData.type
    ];
    sheet.getRange(rowNumber, 1, 1, 15).setValues([rowData]);
    return { success: true, message: "บันทึกข้อมูลเรียบร้อย", item: { ...formData, id: newId, exportStatus: currentExportStatus } };
  } catch (e) { return { success: false, message: e.toString() }; } 
  finally { lock.releaseLock(); }
}

function deleteData(id) {
  const currentUser = Session.getActiveUser().getEmail();
  if (!isUserAdmin(currentUser)) return { success: false, message: "คุณไม่มีสิทธิ์ลบข้อมูล" };
  const sheet = getSheet(SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat();
    const index = ids.indexOf(id);
    if (index === -1) return { success: false, message: "ไม่พบข้อมูล" };
    sheet.deleteRow(index + 2);
    return { success: true, message: "ลบข้อมูลเรียบร้อยแล้ว" };
  } catch (e) { return { success: false, message: e.toString() }; } 
  finally { lock.releaseLock(); }
}

// *** MASTER EXPORT ***
function exportMasterCSV(ftStatus, filterVal, isPreview) {
  const sheet = getSheet(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return isPreview ? [] : { content: "", count: 0 };

  const range = sheet.getRange(2, 1, lastRow - 1, 14);
  const displayValues = range.getDisplayValues();
  
  let csvContent = ""; 
  let previewData = []; 
  let count = 0;
  const timestamp = "Exported " + getDateStr();
  
  let exportStatusValues = sheet.getRange(2, 14, lastRow - 1, 1).getValues();

  for (let i = 0; i < displayValues.length; i++) {
    const row = displayValues[i];
    
    if (row[13] !== "") continue; 
    if (filterVal && !isDateMatchFilter(row[6], filterVal)) continue; 
    if (row[11] !== ftStatus) continue;

    let idCard = row[3].toString().replace(/'/g, "").replace(/[\r\n]+/g, "").trim();
    
    let codeMap = 0;
    if(ftStatus === 'Verified') codeMap = 1;
    else if(ftStatus === 'Pending Result') codeMap = 2; 
    else if(ftStatus === 'Not Verified') codeMap = 3;
    else if(ftStatus === 'In Progress') codeMap = 4;

    if (isPreview) {
        previewData.push({ code: row[1], name: row[2], idCard: idCard, ftStatus: row[11], mappedCode: codeMap });
    } else {
        csvContent += `"${idCard}",${codeMap}\n`;
        count++;
        exportStatusValues[i][0] = timestamp;
    }
  }

  if (isPreview) return previewData;
  if (count > 0) {
      sheet.getRange(2, 14, lastRow - 1, 1).setValues(exportStatusValues);
  }
  
  return { content: "\uFEFF" + csvContent.trim(), count: count, filename: `Master_Export_${ftStatus}_${getDateStr()}.csv` };
}

// ==========================================
// --- ANNUAL VERIFICATION LOGIC ---
// ==========================================

function getAllProviderOptions() {
  const sheet = getSheet(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  return data.map(row => ({
    code: row[1], name: row[2], idCard: row[3], searchText: `${row[1]} | ${row[2]} | ${row[3]}` 
  })).filter(item => item.code && item.name);
}

// 2. Get Annual Data
function getAnnualData(filterVal) { 
  const sheet = getSheet(ANNUAL_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const currentUser = Session.getActiveUser().getEmail();
  let data = [];

  if (lastRow > 1) {
    const maxCols = sheet.getLastColumn();
    const colsToRead = maxCols < 20 ? maxCols : 20; 
    const values = sheet.getRange(2, 1, lastRow - 1, colsToRead).getDisplayValues(); 
    
    data = values.reduce((acc, row, index) => {
        if (filterVal) {
            if (row.length > 17) {
                if (!isDateMatchFilter(row[17], filterVal)) return acc;
            }
        }
        acc.push({
            rowIndex: index + 2,
            id: row[0] || "",
            refCode: row[1] || "",
            name: row[2] || "",
            group: row[3] || "",
            idCard: row[4] || "",
            birthDate: row[5] || "",
            phone: row[6] || "",
            consentStatus: row[7] || "",
            amount: row[8] || "",
            outstanding: row[9] || "",
            deductionStatus: row[10] || "",
            channel: row[11] || "",
            statusProcess: row[12] || "",
            resultDate: row[13] || "",
            result: row[14] || "",
            lastFollowup: row[15] || "",
            officerEmail: row[16] || "",
            submitDate: row[17] || "",
            note: row[18] || "",
            exportStatus: row[19] || ""
        });
        return acc;
    }, []);
    data.sort((a, b) => parseDateForSort(b.submitDate) - parseDateForSort(a.submitDate));
  }
  const configs = getClientConfig();
  return { currentUser: currentUser, isAdmin: configs.isAdmin, results: configs.results, ftStatuses: configs.ftStatuses, data: data };
}

// 3. Import Annual Data
function importAnnualData(records, selectedDateStr) {
  const sheet = getSheet(ANNUAL_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const lastRow = sheet.getLastRow();
    
    const timeNow = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    const uploader = Session.getActiveUser().getEmail(); 
    const logMessage = `${timeNow} โดย ${uploader}`;
    PropertiesService.getScriptProperties().setProperty('LAST_IMPORT_LOG', logMessage);
    
    let existingMap = new Map(); 
    if (lastRow > 1) {
        const allData = sheet.getRange(2, 1, lastRow - 1, 20).getDisplayValues();
        allData.forEach((row, idx) => {
            const clean = String(row[4]).replace(/'/g, "").trim(); 
            if(clean) {
                existingMap.set(clean, { rowIndex: idx + 2, data: row });
            }
        });
    }

    const year = new Date().getFullYear().toString().substr(-2);
    const prefix = `AN-${year}`;
    let maxSeq = 0;
    if (lastRow > 1) {
       const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
       ids.forEach(id => {
           if(String(id).startsWith(prefix)) {
               const parts = String(id).split('-');
               if(parts.length >= 3) maxSeq = Math.max(maxSeq, parseInt(parts[2]) || 0);
           }
       });
    }

    const newRows = [];
    const updates = []; 
    const addedIds = [];
    const updatedIds = [];
    const submitDateVal = selectedDateStr ? formatDateForSheet(selectedDateStr) : formatDateForSheet(new Date());

    function isDifferent(val1, val2) {
        const s1 = String(val1 || "").replace(/^'/, "").trim();
        const s2 = String(val2 || "").replace(/^'/, "").trim();
        if (s1 === s2) return false;
        const n1 = parseFloat(s1.replace(/,/g, ''));
        const n2 = parseFloat(s2.replace(/,/g, ''));
        if (!isNaN(n1) && !isNaN(n2) && n1 === n2) return false;
        return true;
    }

    records.forEach(rec => {
        const cleanCard = String(rec.idCard).replace(/'/g, "").trim();
        const newUpdateValues = [
            String(rec.refCode || ""), String(rec.name || ""), String(rec.group || ""),
            "'" + cleanCard, formatDateForSheet(rec.birthDate), "'" + String(rec.phone || ""),
            String(rec.consentStatus || ""), String(rec.amount || ""), String(rec.outstanding || ""),
            String(rec.deductionStatus || "")
        ];

        if(existingMap.has(cleanCard)) {
            const existing = existingMap.get(cleanCard);
            const currentValues = existing.data.slice(1, 11);
            let changed = false;
            for(let i=0; i<10; i++) {
                if (isDifferent(currentValues[i], newUpdateValues[i])) {
                    changed = true;
                    break;
                }
            }
            if (changed) {
                updates.push({ row: existing.rowIndex, col: 2, data: [newUpdateValues] });
                updatedIds.push(cleanCard); 
            }
        } else {
            // Insert New
            maxSeq++;
            const newId = `${prefix}-${String(maxSeq).padStart(4, '0')}`;
            const rowDataFull = [
                newId, ...newUpdateValues, 
                rec.channel, 
                "", 
                formatDateForSheet(rec.resultDate), rec.result, formatDateForSheet(rec.lastFollowup),
                rec.officerEmail, submitDateVal, rec.note || "", ""
            ];
            newRows.push(rowDataFull);
            addedIds.push(cleanCard);
        }
    });

    updates.forEach(u => {
        sheet.getRange(u.row, u.col, 1, 10).setValues(u.data);
    });

    if(newRows.length > 0) {
        sheet.getRange(lastRow + 1, 1, newRows.length, 20).setValues(newRows);
    }

    return { success: true, updated: updates.length, added: newRows.length, addedIds: addedIds, updatedIds: updatedIds };

  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function getLastImportLog() {
  return PropertiesService.getScriptProperties().getProperty('LAST_IMPORT_LOG') || "";
}

function saveAnnualData(form) {
  const sheet = getSheet(ANNUAL_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    let rowNumber;
    let newId = form.id;
    const lastRow = sheet.getLastRow();
    let currentExportStatus = "";

    if (form.id) {
       const ids = sheet.getRange(2, 1, lastRow > 1 ? lastRow - 1 : 1, 1).getDisplayValues().flat();
       const index = ids.indexOf(String(form.id));
       if (index === -1) throw new Error("ID not found");
       rowNumber = index + 2;
    } else {
       const year = new Date().getFullYear().toString().substr(-2);
       const prefix = `AN-${year}`;
       let maxSeq = 0;
       if (lastRow > 1) {
           const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
           ids.forEach(id => {
               if(String(id).startsWith(prefix)) {
                   const parts = String(id).split('-');
                   if(parts.length >= 3) maxSeq = Math.max(maxSeq, parseInt(parts[2]) || 0);
               }
           });
       }
       newId = `${prefix}-${String(maxSeq + 1).padStart(4, '0')}`;
       rowNumber = lastRow + 1;
    }

    const rowData = [
        newId, form.refCode, form.name, form.group, "'" + form.idCard,
        formatDateForSheet(form.birthDate), "'" + form.phone,
        form.consentStatus, form.amount, form.outstanding, form.deductionStatus,
        form.channel, form.statusProcess, formatDateForSheet(form.resultDate),
        form.result, formatDateForSheet(form.lastFollowup),
        form.officerEmail, formatDateForSheet(form.submitDate),
        form.note, currentExportStatus
    ];

    sheet.getRange(rowNumber, 1, 1, 20).setValues([rowData]);
    return { success: true, message: "บันทึกข้อมูลเรียบร้อย", item: { ...form, id: newId, exportStatus: currentExportStatus } };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updateAnnualNote(id, newNote) {
  const sheet = getSheet(ANNUAL_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const lastRow = sheet.getLastRow();
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
    const index = ids.indexOf(String(id));
    if (index === -1) return { success: false, message: "ไม่พบข้อมูล" };
    sheet.getRange(index + 2, 19).setValue(newNote); 
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function deleteAnnualData(id) {
    const currentUser = Session.getActiveUser().getEmail();
    if (!isUserAdmin(currentUser)) return { success: false, message: "No Permission" };
    const sheet = getSheet(ANNUAL_SHEET_NAME);
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        const ids = sheet.getRange(2, 1, sheet.getLastRow(), 1).getDisplayValues().flat();
        const index = ids.indexOf(id);
        if(index !== -1) { sheet.deleteRow(index+2); return {success:true}; }
        return {success:false, message: "Not found"};
    } catch(e) { return { success: false, message: e.toString() }; }
    finally { lock.releaseLock(); }
}

function exportAnnualCSV(groupType, filterVal, isPreview) {
  const sheet = getSheet(ANNUAL_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return isPreview ? [] : { content: "", count: 0 };

  const range = sheet.getRange(2, 1, lastRow - 1, 20);
  const displayValues = range.getDisplayValues();
  
  let csvContent = ""; 
  let previewData = []; 
  let count = 0;
  const timestamp = "Exported " + getDateStr();
  let newStatuses = displayValues.map(row => [row[19]]); 

  for (let i = 0; i < displayValues.length; i++) {
    const row = displayValues[i];
    if (row[19] !== "") continue;
    if (filterVal && !isDateMatchFilter(row[17], filterVal)) continue;

    const statusProcess = String(row[12]).trim();
    if (statusProcess !== "ผลตรวจออกแล้ว") continue;

    let idCard = row[4].toString().replace(/'/g, "").replace(/[\r\n]+/g, "").trim();
    let status = row[12].toString().replace(/[\r\n]+/g, "").trim();
    
    let shouldExport = true; 

    if (shouldExport) {
        if (isPreview) {
            previewData.push({ refCode: row[1], name: row[2], idCard: idCard, result: row[14] || status });
        } else {
            csvContent += `"${idCard}",1\n`;
            count++;
            newStatuses[i][0] = timestamp;
        }
    }
  }

  if (isPreview) return previewData;
  if (count > 0) sheet.getRange(2, 20, lastRow - 1, 1).setValues(newStatuses);
  
  return { content: "\uFEFF" + csvContent.trim(), count: count, filename: `Annual_Export_${filterVal || 'ALL'}_${getDateStr()}.csv` };
}

function exportAnnualReport(filterVal) {
  const sheet = getSheet(ANNUAL_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { content: "", count: 0 };

  const displayValues = sheet.getRange(2, 1, lastRow - 1, 20).getDisplayValues();
  let csvContent = "ลำดับ,ID,รหัส,ชื่อ-นามสกุล,กลุ่ม,เลขบัตร,วันเกิด,เบอร์โทร,ยินยอม,ยอดเงิน,ยอดค้าง,สถานะหัก,ช่องทาง,สถานะส่ง,วันรับผล,ผล,ติดตามล่าสุด,จนท.,วันที่ส่ง,หมายเหตุ\n";
  let count = 0;

  displayValues.sort((a, b) => parseDateForSort(a[17]) - parseDateForSort(b[17]));

  for (let i = 0; i < displayValues.length; i++) {
    const row = displayValues[i];
    if (filterVal && !isDateMatchFilter(row[17], filterVal)) continue;

    count++;
    const rowString = [
        count, row[0], row[1], row[2], row[3], `'${row[4]}`, 
        row[5], `'${row[6]}`, row[7], row[8], row[9], row[10],
        row[11], row[12], row[13], row[14], row[15],
        row[16], row[17], row[18]
    ].map(f => {
        let cleanVal = String(f || "").replace(/"/g, '""').replace(/[\r\n]+/g, " ");
        return `"${cleanVal}"`;
    }).join(",");
    
    csvContent += rowString + "\n";
  }
  return { content: "\uFEFF" + csvContent.trim(), count: count, filename: `Annual_Report_${filterVal || 'ALL'}_${getDateStr()}.csv` };
}

// ==========================================
// --- ONBOARD DATA LOGIC ---
// ==========================================

function getOnboardData(filterVal) {
  const sheet = getSheet(ONBOARD_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const currentUser = Session.getActiveUser().getEmail();
  let data = [];

  // Lookup Master Status Logic
  const masterSheet = getSheet(SHEET_NAME);
  const masterLastRow = masterSheet.getLastRow();
  let masterStatusMap = new Map();
  
  if (masterLastRow > 1) {
      const masterData = masterSheet.getRange(2, 1, masterLastRow - 1, 12).getValues();
      masterData.forEach(r => {
          let cleanId = String(r[3]).replace(/'/g, "").trim(); 
          let submitDate = r[6];
          let status = r[11];
          if (cleanId) {
              if (!masterStatusMap.has(cleanId)) {
                  masterStatusMap.set(cleanId, { status: status, date: submitDate });
              } else {
                  let current = masterStatusMap.get(cleanId);
                  let currDate = current.date instanceof Date ? current.date : new Date(0);
                  let newDate = submitDate instanceof Date ? submitDate : new Date(0);
                  if (newDate > currDate) {
                      masterStatusMap.set(cleanId, { status: status, date: submitDate });
                  }
              }
          }
      });
  }

  if (lastRow > 1) {
    const maxCols = sheet.getLastColumn();
    // Ensure we read enough columns (Up to 29 now due to checklists)
    const colsToRead = maxCols < 29 ? maxCols : 29;
    const values = sheet.getRange(2, 1, lastRow - 1, colsToRead).getDisplayValues();

    data = values.reduce((acc, row, index) => {
      // row[1] = Training Date
      if (filterVal && !isDateMatchFilter(row[1], filterVal)) return acc;
      
      let history = [];
      try {
         if (row[16] && row[16].startsWith('[')) {
             history = JSON.parse(row[16]);
         }
      } catch (e) {}

      let cleanObId = String(row[6]).replace(/'/g, "").trim(); 
      let lookup = masterStatusMap.get(cleanObId); 
      let realStatus = "";
      let isExpired = false;

      if (lookup) {
          if (lookup.date instanceof Date) {
             const sixMonthsAgo = new Date();
             sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
             if (lookup.date < sixMonthsAgo) {
                 isExpired = true;
             }
          }
          if (!isExpired) {
             realStatus = lookup.status;
          }
      }

      if (!lookup && !realStatus && row.length > 17 && row[17]) {
          realStatus = row[17]; 
      }
      
      const centerVal = row.length > 18 ? row[18] : "";
      const skipVal = row.length > 19 ? row[19] : "";
      const masterTypeVal = row.length > 20 ? row[20] : "";

      acc.push({
        rowIndex: index + 2,
        id: row[0],
        trainingDate: convertToStandardDate(row[1]),
        maidCode: row[2],
        name: row[3],
        group: row[4],
        phone: row[5],
        idCard: row[6], 
        type: row[7],
        latestFollowup: convertToStandardDate(row[8]), 
        openDate: convertToStandardDate(row[11]),
        callStatus: row[12],
        firstJob: row[13],
        jobId: row[14],
        trainer: row[15],
        history: history, 
        fastTrackStatus: realStatus, 
        center: centerVal, 
        skipFastTrack: skipVal, 
        masterType: masterTypeVal 
      });
      return acc;
    }, []);
    data.sort((a, b) => parseDateForSort(b.trainingDate) - parseDateForSort(a.trainingDate));
  }
  
  const configs = getClientConfig();
  return { 
    currentUser: currentUser, 
    isAdmin: configs.isAdmin, 
    groups: configs.onboardGroups, 
    types: configs.onboardTypes,
    masterTypes: configs.masterTypes, 
    onboardCenters: configs.onboardCenters, 
    data: data 
  };
}

function saveOnboardData(form) {
  const sheet = getSheet(ONBOARD_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    let rowNumber;
    let newId = form.id;
    const lastRow = sheet.getLastRow();

    if (form.id) {
       const ids = sheet.getRange(2, 1, lastRow > 1 ? lastRow - 1 : 1, 1).getDisplayValues().flat();
       const index = ids.indexOf(String(form.id));
       if (index === -1) throw new Error("ID not found");
       rowNumber = index + 2;
    } else {
       const year = new Date().getFullYear().toString().substr(-2);
       const prefix = `OB-${year}`;
       let maxSeq = 0;
       if (lastRow > 1) {
           const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
           ids.forEach(id => {
               if(String(id).startsWith(prefix)) {
                   const parts = String(id).split('-');
                   if(parts.length >= 3) maxSeq = Math.max(maxSeq, parseInt(parts[2]) || 0);
               }
           });
       }
       newId = `${prefix}-${String(maxSeq + 1).padStart(4, '0')}`;
       rowNumber = lastRow + 1;
    }
    
    let currentFastTrackStatus = "";
    if (form.id) {
        currentFastTrackStatus = sheet.getRange(rowNumber, 18).getValue();
    }

    let historyJson = "";
    let latestFollowupDate = "";
    if (form.history && form.history.length > 0) {
        form.history.sort((a, b) => parseDateForSort(b.date) - parseDateForSort(a.date));
        historyJson = JSON.stringify(form.history);
        latestFollowupDate = formatDateForSheet(form.history[0].date); 
    }

    const rowData = [
        newId, 
        formatDateForSheet(form.trainingDate), 
        form.maidCode, 
        form.name, 
        form.group,
        "'" + form.phone, 
        "'" + form.idCard, 
        form.type,
        latestFollowupDate, 
        "", 
        "", 
        formatDateForSheet(form.openDate),
        form.callStatus ? "✓" : "", 
        form.firstJob ? "✓" : "", 
        form.jobId,
        form.trainer,
        historyJson, 
        currentFastTrackStatus,
        form.center, 
        form.skipFastTrack ? "TRUE" : "", 
        form.masterType 
    ];

    sheet.getRange(rowNumber, 1, 1, 21).setValues([rowData]); 
    return { success: true, message: "บันทึกข้อมูลเรียบร้อย" };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function deleteOnboardData(id) {
    const currentUser = Session.getActiveUser().getEmail();
    if (!isUserAdmin(currentUser)) return { success: false, message: "No Permission (Admin only)" };
    
    const sheet = getSheet(ONBOARD_SHEET_NAME);
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        const ids = sheet.getRange(2, 1, sheet.getLastRow(), 1).getDisplayValues().flat();
        const index = ids.indexOf(id);
        if(index !== -1) { sheet.deleteRow(index+2); return {success:true}; }
        return {success:false, message: "Not found"};
    } catch(e) { return { success: false, message: e.toString() }; }
    finally { lock.releaseLock(); }
}

function sendToFastTrack(onboardId) {
    const obSheet = getSheet(ONBOARD_SHEET_NAME);
    const masterSheet = getSheet(SHEET_NAME);
    const lock = LockService.getScriptLock();
    
    try {
        lock.waitLock(10000);
        const obIds = obSheet.getRange(2, 1, obSheet.getLastRow()-1, 1).getDisplayValues().flat();
        const obIndex = obIds.indexOf(onboardId);
        if (obIndex === -1) throw new Error("Onboard ID not found");
        
        const obRowRange = obSheet.getRange(obIndex + 2, 1, 1, 21); 
        const obData = obRowRange.getValues()[0];
        
        const maidCode = obData[2];
        const name = obData[3];
        const phone = obData[5];
        const idCard = obData[6];
        const trainingDate = obData[1];
        const officer = obData[15]; 
        const center = obData[18];  
        const masterType = obData[20]; 
        
        if (!name || !idCard) throw new Error("กรุณาระบุ ชื่อ และ เลขบัตรประชาชน ก่อนส่งตรวจ");

        const masterLastRow = masterSheet.getLastRow();
        let masterId = "";
        
        if (masterLastRow > 1) {
            const masterCheckData = masterSheet.getRange(2, 1, masterLastRow - 1, 7).getValues();
            const cleanIdCard = String(idCard).replace(/'/g, "").trim();
            
            let lastSubmitDate = null;
            let foundDuplicate = false;

            for (let i = 0; i < masterCheckData.length; i++) {
                let rowId = String(masterCheckData[i][3]).replace(/'/g, "").trim();
                if (rowId === cleanIdCard) {
                    foundDuplicate = true;
                    let rowDate = masterCheckData[i][6];
                    if (rowDate instanceof Date) {
                        if (!lastSubmitDate || rowDate > lastSubmitDate) {
                            lastSubmitDate = rowDate;
                        }
                    }
                }
            }

            if (foundDuplicate && lastSubmitDate) {
                const sixMonthsAgo = new Date();
                sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
                if (lastSubmitDate > sixMonthsAgo) {
                      return { success: false, message: `ไม่สามารถส่งตรวจได้: เลขบัตรนี้มีการส่งตรวจแล้วเมื่อ ${formatDateForSheet(lastSubmitDate)} (ยังไม่ครบ 6 เดือน)` };
                }
            }
            
            let maxId = 0;
            const existingIds = masterSheet.getRange(2, 1, masterLastRow - 1, 1).getValues().flat();
            existingIds.forEach(id => { let num = Number(id); if (!isNaN(num) && num > maxId) maxId = num; });
            masterId = (maxId + 1).toString();
        } else {
            masterId = "1";
        }

        const submitDate = new Date(); 
        const masterRow = [
            masterId, maidCode, name, "'" + idCard, "'" + phone, 
            trainingDate, formatDateForSheet(submitDate), officer, center, 
            "", "", "รอตรวจสอบ", "", "", masterType 
        ];

        masterSheet.appendRow(masterRow);
        obSheet.getRange(obIndex + 2, 18).setValue("Sent"); 

        return { success: true, message: "ส่งข้อมูลไปยัง Fast Track เรียบร้อยแล้ว" };

    } catch(e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function exportOnboardReport(filterVal) {
  const sheet = getSheet(ONBOARD_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { content: "", count: 0 };

  const values = sheet.getRange(2, 1, lastRow - 1, 21).getDisplayValues();
  let csvContent = "วันที่อบรม,ID,ชื่อ-นามสกุล,กลุ่ม,ศูนย์,เบอร์โทร,เลขบัตร ปชช.,สถานะ(Onboard),ประเภท(Master),วันเปิดระบบ,ผู้ดูแล,สถานะ FT,ติดตามล่าสุด\n";
  let count = 0;
  
  // Sort by Training Date
  values.sort((a, b) => parseDateForSort(b[1]) - parseDateForSort(a[1]));

  for (let i = 0; i < values.length; i++) {
      const row = values[i];
      // Reuse the same filter logic as display
      if (filterVal && !isDateMatchFilter(row[1], filterVal)) continue;
      
      count++;
      
      // Map columns based on structure
      const trainingDate = row[1];
      const maidCode = row[2];
      const name = row[3];
      const group = row[4];
      const center = row[18];
      const phone = row[5];
      const idCard = row[6];
      const statusOnboard = row[7];
      const masterType = row[20];
      const openDate = row[11];
      const trainer = row[15];
      const ftStatus = row[17];
      const latestFollowup = row[8];

      const rowString = [
          trainingDate, maidCode, name, group, center, `"${phone}"`, `"${idCard}"`, 
          statusOnboard, masterType, openDate, trainer, ftStatus, latestFollowup
      ].map(f => `"${String(f || "").replace(/"/g, '""')}"`).join(",");
      
      csvContent += rowString + "\n";
  }
  
  const timestamp = getDateStr();
  return { content: "\uFEFF" + csvContent, count: count, filename: `Onboard_Report_${filterVal || 'ALL'}_${timestamp}.csv` };
}

// ==========================================
// --- FIRST JOB TRACKING LOGIC (SEPARATE SHEET) ---
// ==========================================

// *** UPDATED: Get First Job Tracking Data (Include Checklist) ***
function getFirstBkData() {
  const onboardSheet = getSheet(ONBOARD_SHEET_NAME);
  const firstBkSheet = getSheet(FIRSTBK_SHEET_NAME);
  const currentUser = Session.getActiveUser().getEmail();
  let data = [];
  
  const now = new Date().getTime(); 

  const lastRow = onboardSheet.getLastRow();
  if (lastRow > 1) {
    const obValues = onboardSheet.getRange(2, 1, lastRow - 1, 21).getDisplayValues();
    
    // Read FirstBk Data Map (Read extended columns up to 32)
    let fbMap = new Map();
    const fbLastRow = firstBkSheet.getLastRow();
    if (fbLastRow > 1) {
        // Read up to col 32 (Check_2_8)
        const fbData = firstBkSheet.getRange(2, 1, fbLastRow - 1, 32).getDisplayValues(); 
        fbData.forEach(r => {
             let history = [];
             try { if (r[21] && r[21].startsWith('[')) history = JSON.parse(r[21]); } catch(e) {}

             fbMap.set(String(r[0]), { 
                 bookingCode: r[5], jobId: r[6], cleanDate: r[7], acceptDate: r[8], status: r[9],
                 // Checklist 1.1 - 1.5
                 c1_1: r[10], c1_2: r[11], c1_3: r[12], c1_4: r[13], c1_5: r[14],
                 advice: r[15], officer: r[16], timestamp: r[17],
                 reviewScore: r[18], customerComment: r[19], problemId: r[20],
                 history: history,
                 // Checklist 1.6 - 1.7
                 c1_6: r[22], c1_7: r[23],
                 // Checklist 2.1 - 2.8
                 c2_1: r[24], c2_2: r[25], c2_3: r[26], c2_4: r[27], 
                 c2_5: r[28], c2_6: r[29], c2_7: r[30], c2_8: r[31]
             });
        });
    }

    data = obValues.reduce((acc, row) => {
       const group = String(row[4]).trim();
       const statusOnboard = String(row[7]).trim();
       const id = String(row[0]);
       
       if (group === 'A') {
          const fbRecord = fbMap.get(id);
          const bookingCode = fbRecord ? fbRecord.bookingCode : "";
          const jobId = fbRecord ? fbRecord.jobId : "";
          const cleanDateStr = fbRecord ? fbRecord.cleanDate : "";
          const status = fbRecord ? fbRecord.status : ""; 

          let cleanTime = 0;
          if (cleanDateStr) cleanTime = parseDateForSort(cleanDateStr);

          let tab = 'before';
          if (cleanTime > 0 && now > cleanTime) {
              tab = 'after';
          }
          if (status === 'Done') tab = 'after';

          let processStatus = "รอรับงานแรก";
          if (bookingCode) {
              if (status === 'PreCallDone') processStatus = "โทรเยี่ยมเรียบร้อย";
              else processStatus = "รอโทรเยี่ยม";
          }
          if (status === 'Done') processStatus = "จบงานสมบูรณ์";
          
          const isPreCallDone = status === 'PreCallDone' || status === 'Done';
          const isPostCallDone = status === 'Done';

          const checklist = fbRecord ? {
              // 1.1-1.5
              c1_1: fbRecord.c1_1, c1_2: fbRecord.c1_2, c1_3: fbRecord.c1_3, c1_4: fbRecord.c1_4, c1_5: fbRecord.c1_5,
              // 1.6-1.7
              c1_6: fbRecord.c1_6, c1_7: fbRecord.c1_7,
              // 2.1-2.8
              c2_1: fbRecord.c2_1, c2_2: fbRecord.c2_2, c2_3: fbRecord.c2_3, c2_4: fbRecord.c2_4,
              c2_5: fbRecord.c2_5, c2_6: fbRecord.c2_6, c2_7: fbRecord.c2_7, c2_8: fbRecord.c2_8,
              
              advice: fbRecord.advice,
              officer: fbRecord.officer,
              timestamp: fbRecord.timestamp,
              reviewScore: fbRecord.reviewScore,
              customerComment: fbRecord.customerComment,
              problemId: fbRecord.problemId
          } : null;

          if (statusOnboard === 'เปิดระบบ' || bookingCode) {
             acc.push({
                id: id, maidCode: row[2], name: row[3], phone: row[5], center: row[18],
                bookingCode: bookingCode, jobId: jobId, cleanDate: cleanDateStr,
                processStatus: processStatus,
                isPreCallDone: isPreCallDone, isPostCallDone: isPostCallDone,
                tab: tab, checklist: checklist, history: fbRecord ? fbRecord.history : []
             });
          }
       }
       return acc;
    }, []);
  }

  return { currentUser: currentUser, data: data };
}

// *** Save Job Details (Step 1 - Assign) ***
function saveFirstJobDetails(form) {
  const firstBkSheet = getSheet(FIRSTBK_SHEET_NAME);
  const onboardSheet = getSheet(ONBOARD_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    // Check if exists in FirstBk
    const lastRow = firstBkSheet.getLastRow();
    let rowIndex = -1;
    
    if (lastRow > 1) {
        const ids = firstBkSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
        const found = ids.indexOf(String(form.id));
        if (found !== -1) rowIndex = found + 2;
    }
    
    const status = "Assigned";
    const timestamp = formatDateForSheet(new Date());
    const logEntry = { date: timestamp, note: `[Assign Job] จ่ายงานแรก: ${form.bookingCode} (Clean: ${form.cleanDate})`, by: form.officer };

    if (rowIndex !== -1) {
        // Update
        firstBkSheet.getRange(rowIndex, 6).setValue(form.bookingCode);
        firstBkSheet.getRange(rowIndex, 7).setValue(form.jobId);
        firstBkSheet.getRange(rowIndex, 8).setValue(form.cleanDate);
        firstBkSheet.getRange(rowIndex, 9).setValue(form.acceptDate);
        
        const currStatus = firstBkSheet.getRange(rowIndex, 10).getValue();
        if(!currStatus) firstBkSheet.getRange(rowIndex, 10).setValue(status);
        firstBkSheet.getRange(rowIndex, 18).setValue(timestamp); 
        
        const historyCell = firstBkSheet.getRange(rowIndex, 22);
        let history = [];
        try { const val = historyCell.getValue(); if (val && String(val).startsWith('[')) history = JSON.parse(val); } catch(e) {}
        history.unshift(logEntry);
        historyCell.setValue(JSON.stringify(history));
    } else {
        // Insert New - Need Maid Details from Onboard
        const obLastRow = onboardSheet.getLastRow();
        const obIds = onboardSheet.getRange(2, 1, obLastRow - 1, 1).getDisplayValues().flat();
        const obIndex = obIds.indexOf(String(form.id));
        if (obIndex === -1) throw new Error("ไม่พบข้อมูลใน Onboard");
        
        const obData = onboardSheet.getRange(obIndex + 2, 1, 1, 21).getValues()[0];
        
        // Initial History
        const historyJson = JSON.stringify([logEntry]);

        const newRow = [
            String(form.id), obData[2], obData[3], "'" + obData[5], obData[18],
            form.bookingCode, form.jobId, form.cleanDate, form.acceptDate, status,
            "", "", "", "", "", "", form.officer, timestamp, "", "", "", historyJson,
            "", "", "", "", "", "", "", "", "", "" // Pad empty for new cols
        ];
        firstBkSheet.appendRow(newRow);
    }
    
    return { success: true, message: "บันทึกข้อมูลงานเรียบร้อย" };
  } catch(e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// *** Save Checklist (Step 2 & 3) ***
function saveFirstJobChecklist(form) {
  const firstBkSheet = getSheet(FIRSTBK_SHEET_NAME);
  const onboardSheet = getSheet(ONBOARD_SHEET_NAME); // To update Master Onboard flag
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const lastRow = firstBkSheet.getLastRow();
    const ids = firstBkSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
    const index = ids.indexOf(String(form.id));
    
    if (index === -1) throw new Error("ไม่พบข้อมูลงาน (ต้องบันทึกงานก่อน)");
    const rowNumber = index + 2;

    const historyCell = firstBkSheet.getRange(rowNumber, 22); // Col 22 = History
    let history = [];
    try {
        const val = historyCell.getValue();
        if (val && String(val).startsWith('[')) history = JSON.parse(val);
    } catch(e) {}

    let logTitle = "";
    let isComplete = false;
    let noteContent = "";

    if (form.type === 'precall') {
        if (form.callResult === 'contacted') {
            logTitle = "[Pre-Call] โทรเยี่ยม (ติดต่อได้/สะดวกคุย)";
            isComplete = true; 
            firstBkSheet.getRange(rowNumber, 10).setValue("PreCallDone");
            
            noteContent = `${logTitle}\n` +
              `1.1 รับงาน: ${form.c1_1}\n` +
              `1.2 โทรยืนยัน: ${form.c1_2}\n` +
              `1.3 ระยะทาง: ${form.c1_3}\n` +
              `1.4 แอพ: ${form.c1_4}\n` +
              `1.5 อุปกรณ์: ${form.c1_5}`;

            // Save Pre-Call Data (Cols 11-15)
            const preData = [form.c1_1, form.c1_2, form.c1_3, form.c1_4, form.c1_5];
            firstBkSheet.getRange(rowNumber, 11, 1, 5).setValues([preData]);
            
        } else {
            const reason = form.callResult || "ไม่ระบุ";
            logTitle = `[Pre-Call] โทรไม่ติด (${reason})`;
            noteContent = logTitle;
            isComplete = false; 
        }
    } else {
        // Post-Call
        logTitle = "[Post-Call] ติดตามหลังจบงาน";
        isComplete = true;
        firstBkSheet.getRange(rowNumber, 10).setValue("Done");
        
        // Also mark Onboard Sheet as First Job Done
        const obLastRow = onboardSheet.getLastRow();
        const obIds = onboardSheet.getRange(2, 1, obLastRow - 1, 1).getDisplayValues().flat();
        const obIndex = obIds.indexOf(String(form.id));
        if (obIndex !== -1) {
            onboardSheet.getRange(obIndex + 2, 14).setValue("✓"); 
        }
        
        noteContent = `${logTitle}\n` +
              `-- ส่วนที่ 1 (ต่อ) --\n` +
              `1.6 ถ่ายรูป: ${form.c1_6}\n` +
              `1.7 AI: ${form.c1_7}\n` +
              `-- ส่วนที่ 2 --\n` +
              `2.1 ประเมิน: ${form.c2_1}\n` +
              `2.7 ปัญหา: ${form.c2_7}\n` +
              `Review: ${form.reviewScore || '-'}`;
              
        // Save Post-Call Data (Cols 23-32)
        const postData = [
            form.c1_6, form.c1_7,
            form.c2_1, form.c2_2, form.c2_3, form.c2_4, form.c2_5, form.c2_6, form.c2_7, form.c2_8
        ];
        // Col 23 starts at W
        firstBkSheet.getRange(rowNumber, 23, 1, 10).setValues([postData]);
        
        // Also save Review, Comment, ProblemID (Cols 19-21)
        firstBkSheet.getRange(rowNumber, 19).setValue(form.reviewScore);
        firstBkSheet.getRange(rowNumber, 20).setValue(form.customerComment);
        firstBkSheet.getRange(rowNumber, 21).setValue(form.problemId);
    }
    
    // Save Advice & Officer (Cols 16-17)
    if(form.advice) {
        firstBkSheet.getRange(rowNumber, 16).setValue(form.advice);
        noteContent += `\nNote: ${form.advice}`;
    }
    firstBkSheet.getRange(rowNumber, 17).setValue(form.officer);
    firstBkSheet.getRange(rowNumber, 18).setValue(formatDateForSheet(new Date()));

    const logEntry = {
        date: formatDateForSheet(new Date()),
        note: noteContent,
        by: form.officer
    };
    
    history.unshift(logEntry);
    historyCell.setValue(JSON.stringify(history));

    return { success: true, message: isComplete ? "บันทึกผลเรียบร้อย" : "บันทึกการโทรเรียบร้อย" };
  } catch(e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// *** Export First Job Report (Updated Columns) ***
function exportFirstBkReport() {
  const sheet = getSheet(FIRSTBK_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { content: "", count: 0 };

  const values = sheet.getRange(2, 1, lastRow - 1, 32).getDisplayValues();
  let csvContent = "รหัสแม่บ้าน,ชื่อ-นามสกุล,ศูนย์,เบอร์โทร,รหัสการจอง,JobID,วันทำความสะอาด,สถานะ," + 
                   "1.1รับงาน,1.2โทรยืนยัน,1.3ระยะทาง,1.4แอพ,1.5อุปกรณ์,1.6ถ่ายรูป,1.7AI," +
                   "2.1ประเมิน,2.2รีโนเวท,2.3เกินขอบเขต,2.4จัดการเกิน,2.5สัตว์เลี้ยง,2.6อุปกรณ์หน้างาน,2.7ปัญหา,2.8อื่นๆ," + 
                   "คำแนะนำ,ผู้ติดตาม,รีวิว,คอมเม้น,รหัสปัญหา\n";
  let count = 0;

  for (let i = 0; i < values.length; i++) {
      const row = values[i];
      count++;
      
      const rowString = [
          row[1], row[2], row[4], `"${row[3]}"`, row[5], row[6], row[7], row[9],
          row[10], row[11], row[12], row[13], row[14], row[22], row[23], // 1.x
          row[24], row[25], row[26], row[27], row[28], row[29], row[30], row[31], // 2.x
          row[15], row[16], row[18], row[19], row[20] // Info
      ].map(f => `"${String(f || "").replace(/"/g, '""')}"`).join(",");
      
      csvContent += rowString + "\n";
  }
  
  return { content: "\uFEFF" + csvContent, count: count, filename: `FirstJob_Report_${getDateStr()}.csv` };
}

// --- Utilities ---
function parseDateForSort(dateStr) {
  if (!dateStr) return 0;
  dateStr = String(dateStr).trim();
  
  // Support DD/MM/YYYY HH:mm
  if (dateStr.match(/^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}/)) {
      const [dPart, tPart] = dateStr.split(' ');
      const [d, m, y] = dPart.split('/').map(Number);
      const [hr, min] = tPart.split(':').map(Number);
      let year = y > 2400 ? y - 543 : y;
      return new Date(year, m - 1, d, hr, min).getTime();
  }
  
  if (dateStr.includes(' ')) dateStr = dateStr.split(' ')[0];
  
  // YYYY-MM-DD
  if (dateStr.match(/^\d{4}-\d{2}-\d{2}/)) {
     const parts = dateStr.split('-');
     return new Date(parts[0], parts[1]-1, parts[2]).getTime();
  }
  
  let parts = dateStr.split(/[-/]/);
  if (parts.length === 3) {
      if (parts[0].length === 4) {
           let y = parseInt(parts[0]), m = parseInt(parts[1]) - 1, d = parseInt(parts[2]);
           if (y > 2400) y -= 543;
           return new Date(y, m, d).getTime();
      }
      let d = parseInt(parts[0]), m = parseInt(parts[1]) - 1, y = parseInt(parts[2]);
      if (y > 2400) y -= 543;
      return new Date(y, m, d).getTime();
  }
  return 0;
}
function formatDateForSheet(dateStr) {
  if (!dateStr) return "";
  if (Object.prototype.toString.call(dateStr) === '[object Date]') {
       let d = dateStr.getDate().toString().padStart(2, '0');
       let m = (dateStr.getMonth()+1).toString().padStart(2, '0');
       let y = dateStr.getFullYear();
       return `${d}/${m}/${y > 2400 ? y-543 : y}`;
  }
  if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) { const [year, month, day] = dateStr.split('-'); let y = parseInt(year); if (y > 2400) y -= 543; return `${day}/${month}/${y}`; } 
  return dateStr; 
}
function parseYearFromDate(dateStr) {
    if(!dateStr) return "";
    let parts = String(dateStr).split(/[-/]/);
    let y;
    if (parts.length === 3) {
        if(parts[0].length === 4) y = parseInt(parts[0]); else if(parts[2].length === 4) y = parseInt(parts[2]);
    }
    if(y) return y > 2400 ? y : y + 543; 
    return "";
}
function getDateStr() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm"); }

// *** Generic Date Matcher ***
function isDateMatchFilter(dateStr, filterVal) {
  if (!dateStr || !filterVal) return false;
  dateStr = String(dateStr);
  
  if (dateStr.includes(' ')) dateStr = dateStr.split(' ')[0];

  let separator = null;
  if (filterVal.includes(" to ")) separator = " to ";
  else if (filterVal.includes(" ถึง ")) separator = " ถึง ";
  else if (filterVal.includes(" - ")) separator = " - ";

  if (separator) {
      const [startStr, endStr] = filterVal.split(separator);
      const rowTime = parseDateForSort(dateStr);
      const startTime = parseDateForSort(startStr);
      const endTime = parseDateForSort(endStr);
      
      if (rowTime === 0 || startTime === 0 || endTime === 0) return false;
      return rowTime >= startTime && rowTime <= endTime;
  }
  
  if (filterVal.match(/^\d{4}-\d{2}-\d{2}$/)) {
      const rowTime = parseDateForSort(dateStr);
      const filterTime = parseDateForSort(filterVal);
      if (rowTime === 0) return false; 
      return rowTime === filterTime;
  }
  
  return isDateInMonth(dateStr, filterVal);
}

function isDateInMonth(dateStr, filter) {
  if (!dateStr || !filter) return false;
  dateStr = String(dateStr);
  
  if (dateStr.includes(' ')) {
      dateStr = dateStr.split(' ')[0];
  }

  let filterYear, filterMonth;
  if (filter.indexOf('-') > -1) { [filterYear, filterMonth] = filter.split('-'); } else { filterYear = filter; }
  
  if (dateStr.startsWith(filter)) return true;
  
  let parts = dateStr.split(/[-/]/); 
  if (parts.length === 3) {
      let y, m;
      if (parts[0].length === 4) { y = parts[0]; m = parts[1]; } 
      else if (parts[2].length === 4) { y = parts[2]; m = parts[1]; } 
      
      if (y && m) {
        if (parseInt(y) > 2400) y = (parseInt(y) - 543).toString();
        m = m.toString().padStart(2, '0');
        
        if (filterMonth) return y === filterYear && m === filterMonth; 
        else return y === filterYear;
      }
  }
  return false;
}

function convertToStandardDate(dateStr) {
    if (!dateStr) return "";
    let timestamp = parseDateForSort(dateStr);
    if (timestamp === 0) return dateStr; 
    let d = new Date(timestamp);
    let y = d.getFullYear();
    let m = (d.getMonth() + 1).toString().padStart(2, '0');
    let day = d.getDate().toString().padStart(2, '0');
    return `${y}-${m}-${day}`;
}