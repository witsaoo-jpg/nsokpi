// ============================================================
// KPI Dashboard System - Code.gs (Backend API)
// กลุ่มภารกิจด้านการพยาบาล โรงพยาบาลชลบุรี
// ============================================================

const SPREADSHEET_ID = '1INQVkuKsRefJP5_y1LJQvcI6TgqlAfoOC63I8bVQ8hY';
const SHEETS = {
  USERS: 'Users',
  UNITS: 'Units',
  KPI_MASTER: 'KPI_Master',
  KPI_DATA: 'KPI_Data_Entry',
  AUDIT: 'Audit_Trail'
};

// ============================================================
// ENTRY POINT
// ============================================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('KPI Dashboard - กลุ่มภารกิจด้านการพยาบาล รพ.ชลบุรี')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const data = payload.data || {};
    const token = payload.token || '';

    // Public actions (no auth required)
    if (action === 'login') return respond(handleLogin(data));

    // Authenticated actions
    const session = validateToken(token);
    if (!session.valid) return respond({ success: false, message: 'กรุณาเข้าสู่ระบบใหม่', code: 401 });

    switch (action) {
      // === DATA READ ===
      case 'getUnits':         return respond(getUnits(session));
      case 'getKPIMaster':     return respond(getKPIMaster(session, data));
      case 'getKPIData':       return respond(getKPIData(session, data));
      case 'getDashboardData': return respond(getDashboardData(session, data));
      case 'getUsers':         return respond(getUsers(session));

      // === DATA WRITE ===
      case 'saveBatchKPI':     return respond(saveBatchKPI(session, data));
      case 'saveKPI':          return respond(saveKPIRecord(session, data));
      case 'deleteKPI':        return respond(deleteKPIRecord(session, data));

      // === ADMIN: USER MANAGEMENT ===
      case 'createUser':       return respond(createUser(session, data));
      case 'updateUser':       return respond(updateUser(session, data));
      case 'deleteUser':       return respond(deleteUser(session, data));
      case 'resetPassword':    return respond(resetPassword(session, data));

      // === ADMIN: UNIT MANAGEMENT ===
      case 'createUnit':       return respond(createUnit(session, data));
      case 'updateUnit':       return respond(updateUnit(session, data));
      case 'deleteUnit':       return respond(deleteUnit(session, data));

      // === ADMIN: KPI MASTER ===
      case 'createKPIMaster':  return respond(createKPIMaster(session, data));
      case 'updateKPIMaster':  return respond(updateKPIMaster(session, data));
      case 'deleteKPIMaster':  return respond(deleteKPIMaster(session, data));

      // === USER ===
      case 'changePassword':   return respond(changePassword(session, data));

      case 'getAuditLog':      return respond(getAuditLog(session, data));
      case 'debugInfo':        return respond(debugInfo(session));
      default: return respond({ success: false, message: 'Unknown action' });
    }
  } catch (err) {
    logError(err);
    return respond({ success: false, message: 'Server Error: ' + err.message });
  }
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// AUTHENTICATION
// ============================================================
function handleLogin(data) {
  const { username, password } = data;
  if (!username || !password) return { success: false, message: 'กรุณากรอก Username และรหัสผ่าน' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];

  const colMap = {};
  headers.forEach((h, i) => colMap[h] = i);

  const pwHash = hashSHA256(password);
  // Normalize input to string (รองรับ Username ที่เป็นตัวเลข เช่น Unit_ID = 22)
  const usernameStr = String(username).trim();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    // Convert sheet value to string ก่อนเปรียบเทียบ (แก้ปัญหา number vs string)
    const rowUsername = String(row[colMap['Username']] ?? '').trim();
    const rowPwHash   = String(row[colMap['Password_Hash']] ?? '').trim();
    if (rowUsername === usernameStr && rowPwHash === pwHash) {
      const role   = String(row[colMap['Role']] ?? '');
      const unitId = String(row[colMap['Unit_ID']] ?? '');
      const userId = String(row[colMap['User_ID']] ?? '');

      // Get unit/group info
      const unitInfo = getUnitInfo(unitId);
      const token = generateToken(userId, role, unitId);

      writeAudit(userId, 'LOGIN', 'Users', username);

      return {
        success: true,
        token,
        user: {
          userId,
          username,
          role,
          unitId,
          unitName: unitInfo.unitName,
          groupId: unitInfo.groupId,
          groupName: unitInfo.groupName
        }
      };
    }
  }
  return { success: false, message: 'Username หรือรหัสผ่านไม่ถูกต้อง' };
}

function generateToken(userId, role, unitId) {
  const payload = { userId, role, unitId, exp: Date.now() + 8 * 3600 * 1000 };
  return Utilities.base64Encode(JSON.stringify(payload));
}

function validateToken(token) {
  try {
    const decoded = JSON.parse(Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString());
    if (decoded.exp < Date.now()) return { valid: false };
    return { valid: true, ...decoded };
  } catch (e) {
    return { valid: false };
  }
}

function hashSHA256(input) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

// ============================================================
// UNIT DATA
// ============================================================
function getUnits(session) {
  const rows = getSheetData(SHEETS.UNITS);
  if (!rows) return { success: false, message: 'ไม่พบข้อมูล Units' };

  let units = rows.data;

  // Filter units by role
  if (session.role === 'User') {
    // User: เห็นเฉพาะหน่วยตัวเอง
    const sid = String(session.unitId ?? '').trim();
    units = units.filter(u => String(u.Unit_ID ?? '').trim() === sid);
  } else if (session.role === 'Head') {
    // Head: เห็นทุกวอร์ดในกลุ่มตัวเอง
    const gid = String(session.unitId ?? '').trim();
    units = units.filter(u => String(u.Group_ID ?? '').trim() === gid);
  } else if (session.role === 'Spec') {
    // Spec: เห็นเฉพาะหน่วยตัวเอง
    const sid = String(session.unitId ?? '').trim();
    units = units.filter(u => String(u.Unit_ID ?? '').trim() === sid);
  }

  // Build groups map
  const groupsMap = {};
  units.forEach(u => {
    if (!groupsMap[u.Group_ID]) {
      groupsMap[u.Group_ID] = { groupId: u.Group_ID, groupName: u.Group_Name, units: [] };
    }
    groupsMap[u.Group_ID].units.push({ unitId: u.Unit_ID, unitName: u.Unit_Name });
  });

  return { success: true, units, groups: Object.values(groupsMap) };
}

function getUnitInfo(unitId) {
  const rows = getSheetData(SHEETS.UNITS);
  if (!rows) return { unitName: '', groupId: '', groupName: '' };
  // String comparison รองรับ Unit_ID ที่เป็นตัวเลขใน Sheet
  const uid = String(unitId ?? '').trim();
  const unit = rows.data.find(u => String(u.Unit_ID ?? '').trim() === uid);
  if (!unit) return { unitName: String(unitId), groupId: '', groupName: '' };
  return { unitName: unit.Unit_Name, groupId: unit.Group_ID, groupName: unit.Group_Name };
}

function createUnit(session, data) {
  if (session.role !== 'Admin') return { success: false, message: 'Permission denied' };
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.UNITS);
  sheet.appendRow([data.Group_ID, data.Group_Name, data.Unit_ID, data.Unit_Name]);
  writeAudit(session.userId, 'CREATE', 'Units', data.Unit_ID);
  return { success: true, message: 'เพิ่มหน่วยงานสำเร็จ' };
}

function updateUnit(session, data) {
  if (session.role !== 'Admin') return { success: false, message: 'Permission denied' };
  const result = updateRowWhere(SHEETS.UNITS, 'Unit_ID', data.Unit_ID, data);
  if (result) {
    writeAudit(session.userId, 'UPDATE', 'Units', data.Unit_ID);
    return { success: true, message: 'อัปเดตสำเร็จ' };
  }
  return { success: false, message: 'ไม่พบข้อมูล' };
}

function deleteUnit(session, data) {
  if (session.role !== 'Admin') return { success: false, message: 'Permission denied' };
  const result = deleteRowWhere(SHEETS.UNITS, 'Unit_ID', data.unitId);
  if (result) {
    writeAudit(session.userId, 'DELETE', 'Units', data.unitId);
    return { success: true, message: 'ลบสำเร็จ' };
  }
  return { success: false, message: 'ไม่พบข้อมูล' };
}

// ============================================================
// KPI MASTER
// ============================================================
function getKPIMaster(session, data) {
  const rows = getSheetData(SHEETS.KPI_MASTER);
  if (!rows) return { success: true, kpis: [] };

  const str = v => String(v ?? '').trim();
  // common_KPI/common_PKI/ALL/ว่าง = KPI ที่ใช้ร่วมกันทุกหน่วยงาน
  const isCommon = uid => {
    const u = uid.toLowerCase();
    return !uid || u === '' || u.startsWith('common') || u === 'all';
  };

  let kpis = rows.data.filter(k => str(k.Status) !== 'Inactive');

  if (session.role === 'User') {
    // User (วอร์ด): เห็นเฉพาะ KPI หน่วยตัวเอง + กลุ่มงาน + common
    const unitInfo = getUnitInfo(session.unitId);
    const sid = str(session.unitId);
    const gid = str(unitInfo.groupId);
    kpis = kpis.filter(k => {
      const kid = str(k.Unit_ID);
      return kid === sid || (gid && kid === gid) || isCommon(kid);
    });
  } else if (session.role === 'Head') {
    // Head: เห็นเฉพาะ KPI กลุ่มตัวเอง + วอร์ดในกลุ่ม ไม่เห็น common
    const gid = str(session.unitId);
    const unitsInGroup = getSheetData(SHEETS.UNITS);
    const groupUnitIds = unitsInGroup
      ? unitsInGroup.data.filter(u => str(u.Group_ID) === gid).map(u => str(u.Unit_ID))
      : [];
    kpis = kpis.filter(k => {
      const kid = str(k.Unit_ID);
      return kid === gid || groupUnitIds.includes(kid);
    });
  } else if (session.role === 'Spec') {
    // Spec: เห็นเฉพาะ KPI ที่ Unit_ID ตรงกับตัวเองเท่านั้น — ไม่เห็น common, ไม่เห็นกลุ่ม
    const sid = str(session.unitId);
    kpis = kpis.filter(k => str(k.Unit_ID) === sid);
  } else if (data.unitId && str(data.unitId)) {
    const uid = str(data.unitId);
    const gid = str(data.groupId || '');
    kpis = kpis.filter(k => {
      const kid = str(k.Unit_ID);
      return kid === uid || (gid && kid === gid) || isCommon(kid);
    });
  }

  return { success: true, kpis };
}

function createKPIMaster(session, data) {
  if (session.role !== 'Admin') return { success: false, message: 'Permission denied' };
  const kpiId = 'KPI' + Date.now();
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.KPI_MASTER);
  sheet.appendRow([kpiId, data.KPI_Name, data.KPI_Category, data.Unit_ID, data.Target, data.Status || 'Active', data.Calc_Type]);
  writeAudit(session.userId, 'CREATE', 'KPI_Master', kpiId);
  return { success: true, message: 'สร้าง KPI สำเร็จ', kpiId };
}

function updateKPIMaster(session, data) {
  if (session.role !== 'Admin') return { success: false, message: 'Permission denied' };
  const result = updateRowWhere(SHEETS.KPI_MASTER, 'KPI_ID', data.KPI_ID, data);
  if (result) {
    writeAudit(session.userId, 'UPDATE', 'KPI_Master', data.KPI_ID);
    return { success: true, message: 'อัปเดต KPI สำเร็จ' };
  }
  return { success: false, message: 'ไม่พบ KPI' };
}

function deleteKPIMaster(session, data) {
  if (session.role !== 'Admin') return { success: false, message: 'Permission denied' };
  const result = deleteRowWhere(SHEETS.KPI_MASTER, 'KPI_ID', data.kpiId);
  if (result) {
    writeAudit(session.userId, 'DELETE', 'KPI_Master', data.kpiId);
    return { success: true, message: 'ลบ KPI สำเร็จ' };
  }
  return { success: false, message: 'ไม่พบ KPI' };
}

// ============================================================
// KPI DATA ENTRY
// ============================================================
function getKPIData(session, data) {
  const { unitId, monthYear } = data;
  const rows = getSheetData(SHEETS.KPI_DATA);
  if (!rows) return { success: true, records: [] };

  const str = v => String(v ?? '').trim();
  let records = rows.data;

  // Role-based filter
  if (session.role === 'User') {
    // User: เห็นเฉพาะหน่วยงานตนเอง
    const sid = str(session.unitId);
    records = records.filter(r => str(r.Unit_ID) === sid);
  } else if (session.role === 'Head') {
    // Head: เห็นข้อมูลทุกวอร์ดในกลุ่มตัวเอง
    const gid = str(session.unitId);
    const unitsInGroup = getSheetData(SHEETS.UNITS);
    const groupUnitIds = unitsInGroup
      ? unitsInGroup.data.filter(u => str(u.Group_ID) === gid).map(u => str(u.Unit_ID))
      : [];
    records = records.filter(r => groupUnitIds.includes(str(r.Unit_ID)));
  } else if (session.role === 'Spec') {
    // Spec: เห็นเฉพาะข้อมูลหน่วยตัวเอง
    const sid = str(session.unitId);
    records = records.filter(r => str(r.Unit_ID) === sid);
  } else if (unitId && str(unitId)) {
    const uid = str(unitId);
    records = records.filter(r => str(r.Unit_ID) === uid);
  }

  if (monthYear && str(monthYear)) {
    const my = str(monthYear);
    records = records.filter(r => str(r.Month_Year) === my);
  }

  return { success: true, records };
}

function saveBatchKPI(session, data) {
  const { monthYear, unitId, entries } = data;
  if (!entries || !Array.isArray(entries)) return { success: false, message: 'ข้อมูลไม่ถูกต้อง' };

  // Normalize helper — แก้ปัญหา number vs string
  const s = v => String(v ?? '').trim();

  // Permission check (normalize ก่อนเปรียบเทียบ)
  if (session.role === 'User' && s(unitId) !== s(session.unitId)) {
    return { success: false, message: 'ไม่มีสิทธิ์บันทึกข้อมูลหน่วยงานอื่น' };
  }
  if (session.role === 'Head') {
    // Head บันทึกได้เฉพาะวอร์ดในกลุ่มตัวเอง
    const unitsCheck = getSheetData(SHEETS.UNITS);
    const groupUnitIds = unitsCheck
      ? unitsCheck.data.filter(u => s(u.Group_ID) === s(session.unitId)).map(u => s(u.Unit_ID))
      : [];
    if (!groupUnitIds.includes(s(unitId))) {
      return { success: false, message: 'ไม่มีสิทธิ์บันทึกข้อมูลวอร์ดนอกกลุ่มงานของท่าน' };
    }
  }
  if (session.role === 'Spec') {
    // Spec บันทึกได้เฉพาะหน่วยตัวเองเท่านั้น
    if (s(unitId) !== s(session.unitId)) {
      return { success: false, message: 'ไม่มีสิทธิ์บันทึกข้อมูลหน่วยงานอื่น' };
    }
  }

  const ss2 = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss2.getSheetByName(SHEETS.KPI_DATA);
  // โหลดข้อมูลทั้งหมดครั้งเดียว (batch-read)
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const colMap = {};
  headers.forEach((h, i) => colMap[h] = i);

  const unitIdStr    = s(unitId);
  const monthYearStr = s(monthYear);

  let saved = 0, updated = 0;

  entries.forEach(entry => {
    const { kpiId, resultValue, numerator, denominator } = entry;
    const kpiIdStr = s(kpiId);

    // หา existing row ด้วย String comparison ทุกจุด
    let foundRow = -1;
    for (let i = 1; i < allData.length; i++) {
      if (s(allData[i][colMap['Month_Year']]) === monthYearStr &&
          s(allData[i][colMap['Unit_ID']])    === unitIdStr    &&
          s(allData[i][colMap['KPI_ID']])     === kpiIdStr) {
        foundRow = i;
        break;
      }
    }

    if (foundRow >= 0) {
      // Update existing row
      sheet.getRange(foundRow + 1, colMap['Result_Value'] + 1).setValue(resultValue ?? '');
      sheet.getRange(foundRow + 1, colMap['Numerator']    + 1).setValue(numerator   ?? '');
      sheet.getRange(foundRow + 1, colMap['Denominator']  + 1).setValue(denominator ?? '');
      sheet.getRange(foundRow + 1, colMap['Updated_By']   + 1).setValue(session.userId);
      sheet.getRange(foundRow + 1, colMap['Last_Updated'] + 1).setValue(new Date());
      updated++;
    } else {
      // Insert new row
      const recordId = 'REC' + Date.now() + '_' + Math.random().toString(36).substr(2, 4);
      const newRow = headers.map(h => ({
        'Record_ID':    recordId,
        'Timestamp':    new Date(),
        'Month_Year':   monthYearStr,
        'Unit_ID':      unitIdStr,
        'KPI_ID':       kpiIdStr,
        'Result_Value': resultValue ?? '',
        'Numerator':    numerator   ?? '',
        'Denominator':  denominator ?? '',
        'Recorded_By':  session.userId,
        'Updated_By':   session.userId,
        'Last_Updated': new Date()
      }[h] ?? ''));
      sheet.appendRow(newRow);
      saved++;
    }
  });

  writeAudit(session.userId, 'BATCH_SAVE', 'KPI_Data_Entry', `${unitIdStr}/${monthYearStr}`);
  return { success: true, message: `บันทึกสำเร็จ: เพิ่มใหม่ ${saved} รายการ, อัปเดต ${updated} รายการ` };
}

function deleteKPIRecord(session, data) {
  if (session.role === 'User') return { success: false, message: 'Permission denied' };
  const result = deleteRowWhere(SHEETS.KPI_DATA, 'Record_ID', data.recordId);
  if (result) {
    writeAudit(session.userId, 'DELETE', 'KPI_Data_Entry', data.recordId);
    return { success: true, message: 'ลบสำเร็จ' };
  }
  return { success: false, message: 'ไม่พบข้อมูล' };
}

// ============================================================
// DASHBOARD DATA
// ============================================================
function getDashboardData(session, data) {
  try {
  const { unitId, groupId, year, monthYear } = data;

  const kpiRows  = getSheetData(SHEETS.KPI_MASTER);
  const dataRows = getSheetData(SHEETS.KPI_DATA);
  if (!kpiRows) return { success: false, message: 'ไม่พบ Sheet KPI_Master — กรุณา setupSheets() ก่อน' };
  if (!dataRows) return { success: false, message: 'ไม่พบ Sheet KPI_Data_Entry — กรุณา setupSheets() ก่อน' };

  let kpis    = kpiRows.data.filter(k => String(k.Status || '').trim() !== 'Inactive');
  let records = dataRows.data;

  // แปลง year (CE) → BE string สำหรับ filter Month_Year
  const yearCE = parseInt(year) || new Date().getFullYear();
  const yearBE = String(yearCE + 543);

  // ── Normalize helper: String comparison ──
  const str = v => String(v ?? '').trim();

  // isCommon helper — KPI ที่ใช้ร่วมกันทุกหน่วย
  const isCommon = uid => { const u=(uid||'').toLowerCase(); return !uid||u===''||u.startsWith('common')||u==='all'; };

  // ── Filter by unit/group ──
  if (session.role === 'User') {
    // User (วอร์ด): เฉพาะหน่วยตัวเอง + กลุ่ม + common
    const sid = str(session.unitId);
    const unitInfo2 = getUnitInfo(session.unitId);
    const gid2 = str(unitInfo2.groupId);
    kpis    = kpis.filter(k => { const kid=str(k.Unit_ID); return kid===sid||(gid2&&kid===gid2)||isCommon(kid); });
    records = records.filter(r => str(r.Unit_ID) === sid);
  } else if (session.role === 'Head') {
    // Head: เฉพาะกลุ่มตัวเอง + วอร์ดในกลุ่ม ไม่เห็น common
    const gid = str(session.unitId);
    const unitsData2 = getSheetData(SHEETS.UNITS);
    const groupUnitIds = unitsData2
      ? unitsData2.data.filter(u => str(u.Group_ID) === gid).map(u => str(u.Unit_ID))
      : [];
    kpis    = kpis.filter(k => { const kid=str(k.Unit_ID); return kid===gid||groupUnitIds.includes(kid); });
    records = records.filter(r => groupUnitIds.includes(str(r.Unit_ID)));
  } else if (session.role === 'Spec') {
    // Spec: เฉพาะ KPI ที่ Unit_ID ตรงกับตัวเองเท่านั้น ไม่เห็น common ไม่เห็นกลุ่ม
    const sid = str(session.unitId);
    kpis    = kpis.filter(k => str(k.Unit_ID) === sid);
    records = records.filter(r => str(r.Unit_ID) === sid);
  } else if (unitId && str(unitId)) {
    const uid = str(unitId);
    kpis    = kpis.filter(k => { const kid=str(k.Unit_ID); return kid===uid||isCommon(kid); });
    records = records.filter(r => str(r.Unit_ID) === uid);
  } else if (groupId && str(groupId)) {
    const gid = str(groupId);
    const unitsData  = getSheetData(SHEETS.UNITS);
    const groupUnits = unitsData.data
      .filter(u => str(u.Group_ID) === gid)
      .map(u => str(u.Unit_ID));
    kpis    = kpis.filter(k => { const kid=str(k.Unit_ID); return kid===gid||groupUnits.includes(kid)||isCommon(kid); });
    records = records.filter(r => groupUnits.includes(str(r.Unit_ID)));
  }

  // ── Filter by time ──
  const allRecordsBeforeTimeFilter = records.length;
  if (monthYear && str(monthYear)) {
    const my = str(monthYear);
    records = records.filter(r => str(r.Month_Year) === my);
  } else {
    // endsWith OR includes — รองรับ format "ม.ค. 2569" (มีช่องว่าง) ด้วย
    const filtered = records.filter(r => {
      const my = str(r.Month_Year);
      return my.endsWith(yearBE) || my.includes(yearBE);
    });
    // ถ้ากรองแล้วได้ 0 → แสดงข้อมูลทั้งหมด (ไม่ filter year)
    records = filtered.length > 0 ? filtered : records;
  }

  // ── Build summary ──
  const summary = kpis.map(kpi => {
    const kid      = str(kpi.KPI_ID);
    const kpiRecs  = records.filter(r => str(r.KPI_ID) === kid);
    let picked;
    if (monthYear && str(monthYear)) {
      picked = kpiRecs.find(r => str(r.Month_Year) === str(monthYear));
    } else {
      picked = kpiRecs.sort((a, b) =>
        str(b.Month_Year) > str(a.Month_Year) ? 1 : -1
      )[0];
    }
    return {
      kpiId:       kpi.KPI_ID,
      kpiName:     kpi.KPI_Name,
      category:    kpi.KPI_Category,
      target:      kpi.Target,
      calcType:    kpi.Calc_Type,
      latestValue: picked ? str(picked.Result_Value) || null : null,
      latestMonth: picked ? str(picked.Month_Year)   : null
    };
  });

  // ── Build trend (12 months, ใช้ BE year) ──
  const months = generateMonths(yearBE, 12);

  // เมื่อ filter เฉพาะเดือน trend ดึงจาก all-year records แทน
  let trendRecords = records;
  if (monthYear && str(monthYear)) {
    // ต้อง reload all-year records สำหรับ trend
    const allData = getSheetData(SHEETS.KPI_DATA);
    let ar = allData ? allData.data : [];
    if (session.role === 'User') {
      const sid = str(session.unitId);
      ar = ar.filter(r => str(r.Unit_ID) === sid);
    } else if (unitId && str(unitId)) {
      ar = ar.filter(r => str(r.Unit_ID) === str(unitId));
    }
    trendRecords = ar.filter(r => str(r.Month_Year).endsWith(yearBE));
  }

  const trend = kpis.slice(0, 6).map(kpi => {
    const kid     = str(kpi.KPI_ID);
    const kpiRecs = trendRecords.filter(r => str(r.KPI_ID) === kid);
    return {
      kpiId:   kpi.KPI_ID,
      kpiName: kpi.KPI_Name,
      target:  kpi.Target,
      calcType: kpi.Calc_Type,
      data: months.map(m => {
        const rec = kpiRecs.find(r => str(r.Month_Year) === m);
        return { month: m, value: rec ? (parseFloat(rec.Result_Value) || null) : null };
      })
    };
  });

  const debugMeta = {
    role: session.role,
    unitId_used: session.role==='User' ? str(session.unitId) : str(unitId||'(all)'),
    yearCE, yearBE,
    monthYear_filter: str(monthYear||'(ทั้งปี)'),
    kpis_found: kpis.length,
    records_after_filter: records.length,
    summary_count: summary.length,
    summary_with_value: summary.filter(s=>s.latestValue!==null).length,
    months_sample: months.slice(0,3).join(', '),
  };
  Logger.log('[getDashboardData] ' + JSON.stringify(debugMeta));
  return { success: true, summary, trend, months, _debug: debugMeta };

  } catch(e) {
    Logger.log('[getDashboardData ERROR] ' + e.message + ' | stack: ' + e.stack);
    return { success: false, message: 'Server Error: ' + e.message };
  }
}

function generateMonths(yearBE, count) {
  // yearBE คือปี พ.ศ. string เช่น "2568"
  const thMonths = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  return thMonths.map(m => `${m}${yearBE}`);
}

// ============================================================
// USER MANAGEMENT
// ============================================================
function getUsers(session) {
  if (session.role !== 'Admin') return { success: false, message: 'Permission denied' };
  const rows = getSheetData(SHEETS.USERS);
  if (!rows) return { success: true, users: [] };
  // Remove password hash from response
  const users = rows.data.map(u => {
    const { Password_Hash, ...rest } = u;
    return rest;
  });
  return { success: true, users };
}

function createUser(session, data) {
  if (session.role !== 'Admin') return { success: false, message: 'Permission denied' };
  const userId = 'USR' + Date.now();
  const pwHash = hashSHA256(data.password || 'Password1234');
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.USERS);
  sheet.appendRow([userId, data.Username, pwHash, data.Role, data.Unit_ID, new Date()]);
  writeAudit(session.userId, 'CREATE', 'Users', userId);
  return { success: true, message: 'เพิ่มผู้ใช้สำเร็จ', userId };
}

function updateUser(session, data) {
  if (session.role !== 'Admin') return { success: false, message: 'Permission denied' };
  const updateData = { ...data };
  delete updateData.password; // Don't update password here
  const result = updateRowWhere(SHEETS.USERS, 'User_ID', data.User_ID, updateData);
  if (result) {
    writeAudit(session.userId, 'UPDATE', 'Users', data.User_ID);
    return { success: true, message: 'อัปเดตผู้ใช้สำเร็จ' };
  }
  return { success: false, message: 'ไม่พบผู้ใช้' };
}

function deleteUser(session, data) {
  if (session.role !== 'Admin') return { success: false, message: 'Permission denied' };
  const result = deleteRowWhere(SHEETS.USERS, 'User_ID', data.userId);
  if (result) {
    writeAudit(session.userId, 'DELETE', 'Users', data.userId);
    return { success: true, message: 'ลบผู้ใช้สำเร็จ' };
  }
  return { success: false, message: 'ไม่พบผู้ใช้' };
}

function resetPassword(session, data) {
  if (session.role !== 'Admin') return { success: false, message: 'Permission denied' };
  const newPw = data.newPassword || 'Password1234';
  const pwHash = hashSHA256(newPw);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const colMap = {};
  headers.forEach((h, i) => colMap[h] = i);

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][colMap['User_ID']] === data.userId) {
      sheet.getRange(i + 1, colMap['Password_Hash'] + 1).setValue(pwHash);
      writeAudit(session.userId, 'RESET_PASSWORD', 'Users', data.userId);
      return { success: true, message: 'รีเซ็ตรหัสผ่านสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบผู้ใช้' };
}

function changePassword(session, data) {
  const { oldPassword, newPassword } = data;
  const pwHash = hashSHA256(oldPassword);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const colMap = {};
  headers.forEach((h, i) => colMap[h] = i);

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][colMap['User_ID']] === session.userId) {
      if (rows[i][colMap['Password_Hash']] !== pwHash) {
        return { success: false, message: 'รหัสผ่านเดิมไม่ถูกต้อง' };
      }
      sheet.getRange(i + 1, colMap['Password_Hash'] + 1).setValue(hashSHA256(newPassword));
      writeAudit(session.userId, 'CHANGE_PASSWORD', 'Users', session.userId);
      return { success: true, message: 'เปลี่ยนรหัสผ่านสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบผู้ใช้' };
}

// ============================================================
// UTILITY FUNCTIONS
// ============================================================
function getSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return null;
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return { headers: rows[0] || [], data: [] };
    const headers = rows[0];
    const ID_FIELDS = ['User_ID','Username','Unit_ID','Group_ID','KPI_ID','Record_ID'];
    const data = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = ID_FIELDS.includes(h) ? String(row[i] ?? '').trim() : row[i];
      });
      return obj;
    }).filter(row => {
      // กรองเฉพาะ row ที่มีข้อมูลจริง — ใช้ ID field แรกที่ไม่ว่าง
      const idField = ID_FIELDS.find(f => headers.includes(f));
      if (idField) return String(row[idField] ?? '').trim() !== '';
      // fallback: มีค่าใดค่าหนึ่งไม่ว่าง
      return Object.values(row).some(v => v !== '' && v !== null && v !== undefined);
    });
    return { headers, data };
  } catch (e) {
    logError(e);
    return null;
  }
}

function updateRowWhere(sheetName, keyCol, keyVal, updateData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const colMap = {};
  headers.forEach((h, i) => colMap[h] = i);

  const keyColIdx = colMap[keyCol];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][keyColIdx] === keyVal) {
      headers.forEach((h, idx) => {
        if (updateData[h] !== undefined) {
          sheet.getRange(i + 1, idx + 1).setValue(updateData[h]);
        }
      });
      return true;
    }
  }
  return false;
}

function deleteRowWhere(sheetName, keyCol, keyVal) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const keyColIdx = headers.indexOf(keyCol);

  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][keyColIdx] === keyVal) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function writeAudit(userId, action, table, recordId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEETS.AUDIT);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.AUDIT);
      sheet.appendRow(['Timestamp', 'User_ID', 'Action', 'Table', 'Record_ID']);
    }
    sheet.appendRow([new Date(), userId, action, table, recordId]);
  } catch (e) {
    // Audit failure should not crash main operation
  }
}

function logError(err) {
  console.error('[KPI System Error]', err.message, err.stack);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// AUDIT LOG
// ============================================================
function getAuditLog(session, data) {
  if (session.role !== 'Admin') return { success: false, message: 'Admin only' };

  const rows = getSheetData(SHEETS.AUDIT);
  if (!rows || !rows.data.length) return { success: true, logs: [], total: 0 };

  const str = v => String(v ?? '').trim();
  const { filterAction, filterTable, filterUser, limit } = data;
  const maxRows = Math.min(parseInt(limit) || 200, 500);

  // ดึง Username map จาก Users sheet
  const userRows = getSheetData(SHEETS.USERS);
  const userMap = {};
  if (userRows) {
    userRows.data.forEach(u => { userMap[str(u.User_ID)] = str(u.Username); });
  }

  let logs = [...rows.data].reverse(); // newest first

  // Filter
  if (filterAction) logs = logs.filter(r => str(r.Action) === filterAction);
  if (filterTable)  logs = logs.filter(r => str(r.Table)  === filterTable);
  if (filterUser)   logs = logs.filter(r => str(r.User_ID) === filterUser || userMap[str(r.User_ID)] === filterUser);

  const total = logs.length;
  logs = logs.slice(0, maxRows);

  // Enrich with username
  const enriched = logs.map(r => ({
    timestamp: r.Timestamp ? new Date(r.Timestamp).toLocaleString('th-TH') : '—',
    userId:    str(r.User_ID),
    username:  userMap[str(r.User_ID)] || str(r.User_ID),
    action:    str(r.Action),
    table:     str(r.Table),
    recordId:  str(r.Record_ID),
  }));

  // Unique values for filter dropdowns
  const actions = [...new Set(rows.data.map(r => str(r.Action)).filter(Boolean))];
  const tables  = [...new Set(rows.data.map(r => str(r.Table)).filter(Boolean))];
  const users   = [...new Set(rows.data.map(r => userMap[str(r.User_ID)] || str(r.User_ID)).filter(Boolean))];

  return { success: true, logs: enriched, total, actions, tables, users };
}

// ============================================================
// DEBUG UTILITY — ใช้ตรวจสอบข้อมูลใน Sheet
// ============================================================
function debugInfo(session) {
  if (session.role !== 'Admin') return { success: false, message: 'Admin only' };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const result = {};
  ['Users','Units','KPI_Master','KPI_Data_Entry'].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { result[name] = 'NOT FOUND'; return; }
    const data = sheet.getDataRange().getValues();
    result[name] = { rows: data.length - 1, headers: data[0], sample: data.slice(1,3) };
  });
  return { success: true, debug: result };
}

// ============================================================
// SETUP UTILITY — รันครั้งเดียวเพื่อสร้าง Sheets + ข้อมูลจริง
// วิธีใช้: Script Editor → เลือก setupSheets → กด ▶ Run
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // === 1. สร้าง Sheets + Headers ===
  const config = [
    { name: 'Users',         headers: ['User_ID','Username','Password_Hash','Role','Unit_ID','Created_Date'] },
    { name: 'Units',         headers: ['Group_ID','Group_Name','Unit_ID','Unit_Name'] },
    { name: 'KPI_Master',    headers: ['KPI_ID','KPI_Name','KPI_Category','Unit_ID','Target','Status','Calc_Type'] },
    { name: 'KPI_Data_Entry',headers: ['Record_ID','Timestamp','Month_Year','Unit_ID','KPI_ID','Result_Value','Numerator','Denominator','Recorded_By','Updated_By','Last_Updated'] },
    { name: 'Audit_Trail',   headers: ['Timestamp','User_ID','Action','Table','Record_ID'] },
  ];

  config.forEach(({ name, headers }) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      Logger.log('Created sheet: ' + name);
    }
    const firstCell = sheet.getRange(1, 1).getValue();
    if (!firstCell) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#359286')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
      Logger.log('Headers written: ' + name);
    }
  });

  // === 2. นำเข้าข้อมูลกลุ่มงาน/หน่วยงาน (74 หน่วยงาน, 16 กลุ่มงาน) ===
  const UNIT_DATA = [
    ['g01', 'กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม', '13', 'หอผู้ป่วย สก.3'],
    ['g01', 'กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม', '14', 'หอผู้ป่วย สก.4'],
    ['g01', 'กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม', '15', 'หอผู้ป่วย สก.5'],
    ['g01', 'กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม', '16', 'หอผู้ป่วย สก.6'],
    ['g01', 'กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม', '90', 'หอผู้ป่วยสงฆ์อาพาธ'],
    ['g01', 'กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม', '91', 'หอผู้ป่วยสงฆ์พิเศษ 2+3'],
    ['g01', 'กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม', '11', 'หอผู้ป่วยสามัญติดเชื้อ ธารน้ำใจ 1'],
    ['g01', 'กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม', '18', 'หอผู้ป่วยสามัญติดเชื้อ ธารน้ำใจ 2'],
    ['g01', 'กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม', '40', 'หอผู้ป่วยพิเศษอายุรกรรมชลาทรล่าง'],
    ['g01', 'กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม', '17', 'หอผู้ป่วยพิเศษอายุรกรรมชลาทรบน'],
    ['g01', 'กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม', '29', 'หอผู้ป่วย Low Immune ธนจ.4'],
    ['g02', 'กลุ่มงานการพยาบาลผู้ป่วยศัลยกรรม', '22', 'หอผู้ป่วยชลาทิศ 1'],
    ['g02', 'กลุ่มงานการพยาบาลผู้ป่วยศัลยกรรม', '24', 'หอผู้ป่วยชลาทิศ 2'],
    ['g02', 'กลุ่มงานการพยาบาลผู้ป่วยศัลยกรรม', '39', 'หอผู้ป่วยชลาทิศ 3'],
    ['g02', 'กลุ่มงานการพยาบาลผู้ป่วยศัลยกรรม', '25', 'หอผู้ป่วยชลาทิศ 4'],
    ['g02', 'กลุ่มงานการพยาบาลผู้ป่วยศัลยกรรม', '27', 'หอผู้ป่วยพิเศษศัลยกรรม ฉ.7'],
    ['g02', 'กลุ่มงานการพยาบาลผู้ป่วยศัลยกรรม', '28', 'หอผู้ป่วยพิเศษศัลยกรรม ฉ.8'],
    ['g02', 'กลุ่มงานการพยาบาลผู้ป่วยศัลยกรรม', '94', 'หอผู้ป่วยพิเศษศัลยกรรม Ex.9'],
    ['g02', 'กลุ่มงานการพยาบาลผู้ป่วยศัลยกรรม', '23', 'หอผู้ป่วยแผลไหม้'],
    ['g02', 'กลุ่มงานการพยาบาลผู้ป่วยศัลยกรรม', '34', 'หอผู้ป่วยคมีบำบัด'],
    ['g03', 'กลุ่มงานการพยาบาลผู้ป่วยสูติ-นรีเวช', '30', 'หอผู้ป่วยหลังคลอด'],
    ['g03', 'กลุ่มงานการพยาบาลผู้ป่วยสูติ-นรีเวช', '41', 'หอผู้ป่วยนรีเวช ชลารักษ์4'],
    ['g03', 'กลุ่มงานการพยาบาลผู้ป่วยสูติ-นรีเวช', '42', 'หอผู้ป่วยพิเศษนรีเวช ชลารักษ์4'],
    ['g04', 'กลุ่มงานการพยาบาลผู้ป่วยออร์โธปิดิกส์', '81', 'หอผู้ป่วยกระดูกชาย'],
    ['g04', 'กลุ่มงานการพยาบาลผู้ป่วยออร์โธปิดิกส์', '82', 'หอผู้ป่วยศัลยกรรมอุบัติเหตุและกระดูกหญิง'],
    ['g04', 'กลุ่มงานการพยาบาลผู้ป่วยออร์โธปิดิกส์', '44', 'หอผู้ป่วยพิเศษศัลยกรรม Ex.8'],
    ['g05', 'กลุ่มงานการพยาบาลผู้ป่วย โสต ศอ นาสิก จักษุ', '61', 'หอผู้ป่วยสามัญ EENT และศัลยกรรมเด็ก ชว.3'],
    ['g05', 'กลุ่มงานการพยาบาลผู้ป่วย โสต ศอ นาสิก จักษุ', '10', 'หอผู้ป่วยพิเศษ EENT'],
    ['g06', 'กลุ่มงานการพยาบาลผู้ป่วยจิตเวช', '92', 'จิตเวช (ชลาธาร 2) ชาย'],
    ['g06', 'กลุ่มงานการพยาบาลผู้ป่วยจิตเวช', '72', 'จิตเวช (ชลาธาร 3) หญิง'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '20', 'หอผู้ป่วยหนักศัลยกรรม (SICU)'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '12', 'หอผู้ป่วยหนักอายุรกรรม (MICU 1)'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '53', 'หอผู้ป่วยหนักกุมารเวชกรรม (PICU)'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '58', 'หอผู้ป่วยหนักทารกแรกเกิด (NICU 1)'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '59', 'หอผู้ป่วยหนักทารกแรกเกิด (NICU 2)'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '33', 'หอผู้ป่วยหนักโรคหัวใจ ( CCU) )'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '21', 'หอผู้ป่วยหนักโรคหลอดเลือดสมอง ( ชลาธาร 4 )'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '67', 'หอผู้ป่วยหนักโรคติดเชื้อ ( IICU ธารน้ำใจ 3 )'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '37', 'หอผู้ป่วยหนักอุบัติเหตุหลายระบบ (ICU Trauma)'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '21', 'หอผู้ป่วยหนักระบบประสาท (ICU Neuro)'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '38', 'หอผู้ป่วยหนัก CVT'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '75', 'หอผู้ป่วยหนัก CICU (สวนหัวใจ)'],
    ['g07', 'กลุ่มงานการพยาบาลผู้ป่วยหนัก', '36', 'หอผู้ป่วยโรคหลอดเลือดสมอง (ชลาธาร4)'],
    ['g08', 'กลุ่มงานการพยาบาลผู้ป่วยนอก', 'g081', 'OPD อายุรกรรม'],
    ['g08', 'กลุ่มงานการพยาบาลผู้ป่วยนอก', 'g082', 'OPD ศัลยกรรม'],
    ['g08', 'กลุ่มงานการพยาบาลผู้ป่วยนอก', 'g083', 'OPD สูติ-นรีเวช'],
    ['g08', 'กลุ่มงานการพยาบาลผู้ป่วยนอก', 'g084', 'Clinic NCD'],
    ['g08', 'กลุ่มงานการพยาบาลผู้ป่วยนอก', 'g085', 'OPD จิตเวช'],
    ['g08', 'กลุ่มงานการพยาบาลผู้ป่วยนอก', 'g086', 'OPD กระดูก'],
    ['g08', 'กลุ่มงานการพยาบาลผู้ป่วยนอก', 'g087', 'OPD ผิวหนัง'],
    ['g08', 'กลุ่มงานการพยาบาลผู้ป่วยนอก', 'g088', 'OPD GP'],
    ['g08', 'กลุ่มงานการพยาบาลผู้ป่วยนอก', 'g089', 'ห้อง EEG & EKG'],
    ['g09', 'กลุ่มงานการพยาบาลผู้ป่วยอุบัติเหตุ และฉุกเฉิน', 'g091', 'EMS'],
    ['g09', 'กลุ่มงานการพยาบาลผู้ป่วยอุบัติเหตุ และฉุกเฉิน', 'g092', 'ER & Extended ER'],
    ['g09', 'กลุ่มงานการพยาบาลผู้ป่วยอุบัติเหตุ และฉุกเฉิน', 'g093', 'ห้องสังเกตอาการ'],
    ['g09', 'กลุ่มงานการพยาบาลผู้ป่วยอุบัติเหตุ และฉุกเฉิน', 'g094', 'การพยาบาลส่งต่อ'],
    ['g10', 'กลุ่มงานการพยาบาลผู้คลอด', '30', 'ห้องคลอด'],
    ['g11', 'กลุ่มงานการพยาบาลผู้ป่วยห้องผ่าตัด', 'g111', 'ผ่าตัดใหญ่ความเสี่ยงสูง'],
    ['g11', 'กลุ่มงานการพยาบาลผู้ป่วยห้องผ่าตัด', 'g112', 'ผ่าตัดเฉพาะทาง'],
    ['g12', 'กลุ่มงานการพยาบาลวิสัญญี', 'g121', 'วิสัญญีผู้ป่วยความเสี่ยงสูง'],
    ['g12', 'กลุ่มงานการพยาบาลวิสัญญี', 'g122', 'วิสัญญีเฉพาะทาง'],
    ['g12', 'กลุ่มงานการพยาบาลวิสัญญี', 'g123', 'Recovery Room'],
    ['g13', 'กลุ่มงานการพยาบาลผู้ป่วยกุมารเวชกรรม', '52', 'หอผู้ป่วย กุมารเวชกรรม 1'],
    ['g13', 'กลุ่มงานการพยาบาลผู้ป่วยกุมารเวชกรรม', '51', 'หอผู้ป่วย กุมารเวชกรรม 4'],
    ['g13', 'กลุ่มงานการพยาบาลผู้ป่วยกุมารเวชกรรม', '56', 'หอผู้ป่วย กุมาร 5'],
    ['g13', 'กลุ่มงานการพยาบาลผู้ป่วยกุมารเวชกรรม', '64', 'หอผู้ป่วย กุมาร 6'],
    ['g13', 'กลุ่มงานการพยาบาลผู้ป่วยกุมารเวชกรรม', '55', 'หอผู้ป่วย SNB 1'],
    ['g13', 'กลุ่มงานการพยาบาลผู้ป่วยกุมารเวชกรรม', '57', 'หอผู้ป่วย SNB 2'],
    ['g14', 'กลุ่มงานการพยาบาลด้านการควบคุมและป้องกันการติดเชื้อ', 'g141', 'Infection Control'],
    ['g15', 'กลุ่มงานการพยาบาลตรวจรักษาพิเศษ', 'g151', 'ศูนย์โรคหัวใจ'],
    ['g15', 'กลุ่มงานการพยาบาลตรวจรักษาพิเศษ', 'g152', 'งานห้องตรวจพิเศษ'],
    ['g15', 'กลุ่มงานการพยาบาลตรวจรักษาพิเศษ', 'g153', 'หน่วยไตเทียม'],
    ['g15', 'กลุ่มงานการพยาบาลตรวจรักษาพิเศษ', 'g154', 'ODS'],
    ['g16', 'กลุ่มงานวิจัยและพัฒนาการพยาบาล', 'g161', 'งานวิจัยและพัฒนาการพยาบาล'],
  ];

  const unitsSheet = ss.getSheetByName('Units');
  const existingUnits = unitsSheet.getDataRange().getValues();
  const existingUnitIds = existingUnits.slice(1).map(r => String(r[2]));

  let unitsAdded = 0;
  UNIT_DATA.forEach(([gId, gName, uId, uName]) => {
    if (!existingUnitIds.includes(String(uId))) {
      unitsSheet.appendRow([gId, gName, uId, uName]);
      unitsAdded++;
    }
  });
  Logger.log(`Units: ${unitsAdded} เพิ่มใหม่ (ข้ามซ้ำ ${UNIT_DATA.length - unitsAdded} รายการ)`);

  // === 3. สร้าง Admin User เริ่มต้น (admin / admin1234) ===
  const usersSheet = ss.getSheetByName('Users');
  const existingUsers = usersSheet.getDataRange().getValues();
  const hasAdmin = existingUsers.slice(1).some(row => row[1] === 'admin');

  if (!hasAdmin) {
    const pwHash = hashSHA256('admin1234');
    usersSheet.appendRow(['USR001', 'admin', pwHash, 'Admin', '', new Date()]);
    Logger.log('✅ Admin user created: admin / admin1234');
  } else {
    Logger.log('ℹ️ Admin user already exists, skipped.');
  }

  Logger.log('✅ Setup สำเร็จ! ' + SPREADSHEET_ID);
  Logger.log('🔑 Login: admin / admin1234');
  Logger.log('📊 Units loaded: ' + unitsAdded + ' หน่วยงาน');
}