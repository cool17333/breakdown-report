// =====================================================
// Google Apps Script — Breakdown Database
// วิธีติดตั้ง:
//  1. ไปที่ script.google.com → New project
//  2. วางโค้ดนี้ทั้งหมดแทนที่โค้ดเดิม
//  3. ใส่ค่า SPREADSHEET_ID ด้านล่าง
//  4. Deploy → New deployment → Web app
//     - Execute as: Me
//     - Who has access: Anyone
//  5. คัดลอก Web App URL ไปใส่ในหน้า Settings ของ app
// =====================================================

const SPREADSHEET_ID = '1knVTnZf7Ecu-LxOYv4911J-QHwr-PuSgrd44JY_UFqc'; // ID ของ Google Sheet

// รหัสผ่านตามบทบาท (เก็บฝั่ง server — ไม่โผล่ในหน้าเว็บ)
const ROLE_PW = {
  'engineer123456': 'engineer',
  'cpram123456':    'admin',
};

const HEADERS = [
  'วันที่บันทึก',
  'ชื่อเครื่องจักร',
  'โรงงาน',
  'พื้นที่',
  'รหัสเครื่องจักร',
  'สาย / ตำแหน่ง',
  'สถานะ',
  'เวลาเริ่ม Breakdown',
  'เวลาเสร็จสิ้น',
  'Downtime (นาที)',
  'ประเภท Breakdown',
  'ปัญหาที่พบ',
  'อุปกรณ์ที่เกิดปัญหา',
  'Why 1', 'Why 2', 'Why 3', 'Why 4', 'Why 5',
  'มาตรการแก้ไข',
  'มาตรการป้องกัน',
  'อายุมาตรฐาน (Std.)',
  'อายุจริงตอนเสีย',
  'หมายเหตุอายุอะไหล่',
  'อะไหล่ที่ใช้',
  'เลข Tracking',         // index 24
  'ผู้ดำเนินการล่าสุด',   // index 25
  'ผู้รับงาน',            // index 26
  'ผู้ปิดงาน',            // index 27
  'รูปก่อนแก้ไข (ID)',    // index 28
  'รูปหลังแก้ไข (ID)',    // index 29
];

// ============================================================
// POST — บันทึก / อัปเดต
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);

    // อัปโหลดรูปขึ้น Drive (เฉพาะตอนสร้าง/แก้ไข) → เก็บเป็น fileId
    if (data.action === 'create' || data.action === 'update' || !data.action) {
      data.imgBefore = saveImgToDrive(data.imgBefore);
      data.imgAfter  = saveImgToDrive(data.imgAfter);
    }

    // ---- UPDATE existing row ----
    if (data.action === 'update') {
      const sheet = ss.getSheetByName(data.sheetName);
      if (!sheet || !data.rowIndex) throw new Error('Sheet or rowIndex not found');

      // ผู้รับงาน = คนแรกที่รับงาน (ไม่ทับ) / ผู้ปิดงาน = คนที่กดปิดงาน
      const prev = sheet.getRange(data.rowIndex, 1, 1, HEADERS.length).getValues()[0];
      data.acceptedBy = prev[26] || '';
      data.closedBy   = prev[27] || '';
      if (data.status === 'ดำเนินการเสร็จสิ้น') {
        if (!data.closedBy)   data.closedBy   = data.byName || '';   // เซ็ตครั้งแรกที่ปิด (แก้ Why-Why ภายหลังไม่ทับ)
        if (!data.acceptedBy) data.acceptedBy = data.byName || '';
      } else if (data.status === 'กำลังดำเนินการแก้ไข' || data.status === 'รออะไหล่') {
        if (!data.acceptedBy) data.acceptedBy = data.byName || '';
      }

      const whys    = data.whys || [];
      const partsStr = buildPartsStr(data.parts);
      const row = buildRow(data, whys, partsStr, /*keepTimestamp=*/true);
      sheet.getRange(data.rowIndex, 1, 1, row.length).setValues([row]);

      writeLog(ss, data.tracking, 'แก้ไข → ' + (data.status || ''), data.byName, data.status);
      return jsonOut({ success: true, action: 'updated' });
    }

    // ---- SET machines (เขียนทับแท็บ _Machines — มีสำรอง + กันลบทั้งหมด) ----
    if (data.action === 'setMachines') {
      if (ROLE_PW[(data.pw || '').trim()] !== 'admin')
        return jsonOut({ success: false, error: 'ต้องเป็น Admin เท่านั้น' });
      const incoming = (data.machines || []).filter(m => m && String(m.id).trim());
      let sh = ss.getSheetByName('_Machines');
      const existingCount = sh ? Math.max(0, sh.getLastRow() - 1) : 0;

      // กันข้อมูลหาย: ของใหม่ว่าง แต่ของเดิมมี → ปฏิเสธ (เว้นแต่ตั้งใจล้างด้วย force)
      if (incoming.length === 0 && existingCount > 0 && !data.force)
        return jsonOut({ success: false, error: 'รายการว่าง — ยกเลิกเพื่อกันข้อมูลหาย' });

      // สำรองของเดิมไว้ที่ _Machines_bak ก่อนเขียนทับ
      if (sh && sh.getLastRow() > 1) {
        let bak = ss.getSheetByName('_Machines_bak') || ss.insertSheet('_Machines_bak');
        bak.clearContents();
        const cur = sh.getDataRange().getValues();
        bak.getRange(1, 1, cur.length, cur[0].length).setValues(cur);
      }

      if (!sh) sh = ss.insertSheet('_Machines');
      sh.clearContents();
      const header = ['รหัสเครื่องจักร', 'ชื่อเครื่องจักร', 'โรงงาน', 'พื้นที่', 'ไลน์'];
      sh.getRange(1, 1, 1, header.length).setValues([header]);
      const rows = incoming.map(m => [m.id || '', m.name || '', m.factory || '', m.area || '', m.line || '']);
      if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
      return jsonOut({ success: true, count: rows.length, backedUp: existingCount });
    }

    // ---- RESTORE machines จาก _Machines_bak (กู้คืน) ----
    if (data.action === 'restoreMachines') {
      if (ROLE_PW[(data.pw || '').trim()] !== 'admin')
        return jsonOut({ success: false, error: 'ต้องเป็น Admin เท่านั้น' });
      const bak = ss.getSheetByName('_Machines_bak');
      if (!bak || bak.getLastRow() < 2) return jsonOut({ success: false, error: 'ไม่มีข้อมูลสำรอง' });
      let sh = ss.getSheetByName('_Machines') || ss.insertSheet('_Machines');
      sh.clearContents();
      const cur = bak.getDataRange().getValues();
      sh.getRange(1, 1, cur.length, cur[0].length).setValues(cur);
      return jsonOut({ success: true, count: cur.length - 1 });
    }

    // ---- DELETE row (Admin เท่านั้น — เช็ครหัสฝั่ง server) ----
    if (data.action === 'delete') {
      if (ROLE_PW[(data.pw || '').trim()] !== 'admin')
        return jsonOut({ success: false, error: 'ต้องเป็น Admin เท่านั้น' });
      const sheet = ss.getSheetByName(data.sheetName);
      if (!sheet || !data.rowIndex) throw new Error('Sheet or rowIndex not found');
      sheet.deleteRow(data.rowIndex);
      writeLog(ss, data.tracking, 'ลบเอกสาร', data.byName, 'ลบแล้ว');
      return jsonOut({ success: true, action: 'deleted' });
    }

    // ---- CREATE new row ----
    // ล็อกกันเลขรันชนกันเวลาหลายคนแจ้งพร้อมกัน
    const lock = LockService.getScriptLock();
    try { lock.waitLock(15000); } catch (e) {}

    const factoryCode = (data.factory || 'Unknown').replace(/\s+/g, '');
    const now         = new Date();
    const month       = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM');
    const sheetName   = factoryCode + '_' + month;

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const hRange = sheet.getRange(1, 1, 1, HEADERS.length);
      hRange.setValues([HEADERS]);
      hRange.setBackground('#f97316');
      hRange.setFontColor('#ffffff');
      hRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1, HEADERS.length, 120);
      sheet.setColumnWidth(1, 140);
      sheet.setColumnWidth(12, 240);
      sheet.setColumnWidth(13, 200);
    }

    // เลขรันต่อเนื่อง แยกตามโรงงาน+เดือน → BD-CHB{n}-{YYYYMM}-{NNN}
    // seq = จำนวนแถวที่มีอยู่ (header=1 → รายการแรก = 1)
    const seq = sheet.getLastRow();
    const ym  = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMM');
    data.tracking = 'BD-' + factoryToCHB(data.factory) + '-' + ym + '-' + String(seq).padStart(3, '0');

    const whys    = data.whys || [];
    const partsStr = buildPartsStr(data.parts);
    sheet.appendRow(buildRow(data, whys, partsStr, /*keepTimestamp=*/false, now));
    SpreadsheetApp.flush();
    try { lock.releaseLock(); } catch (e) {}

    writeLog(ss, data.tracking, 'แจ้ง Breakdown (สร้างใหม่)', data.byName, data.status);
    return jsonOut({ success: true, sheet: sheetName, tracking: data.tracking });

  } catch (err) {
    return jsonOut({ success: false, error: err.toString() });
  }
}

function buildRow(data, whys, partsStr, keepTimestamp, now) {
  const ts = keepTimestamp
    ? (data.timestamp || '')
    : Utilities.formatDate(now || new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  return [
    ts,
    data.machineName   || '',
    data.factory       || '',
    data.area          || '',
    data.machineId     || '',
    data.line          || '',
    data.status        || '',
    data.bdStart       || '',
    data.bdEnd         || '',
    Number(data.downtimeMin) || 0,
    data.bdType        || '',
    data.problem       || '',
    data.device        || '',
    whys[0]||'', whys[1]||'', whys[2]||'', whys[3]||'', whys[4]||'',
    data.corrective    || '',
    data.preventive    || '',
    data.stdLife       || '',
    data.actualLife    || '',
    data.lifeNote      || '',
    partsStr,
    data.tracking      || '',
    data.byName        || '',
    data.acceptedBy    || '',
    data.closedBy      || '',
    data.imgBefore     || '',
    data.imgAfter      || '',
  ];
}

// รันฟังก์ชันนี้ 1 ครั้งในตัว editor เพื่ออนุญาตสิทธิ์ Google Drive (แก้ error DriveApp ไม่ได้รับอนุญาต)
function authorizeDrive() {
  DriveApp.getRootFolder().getName();   // กด Run → Allow สิทธิ์ Drive
  return 'OK — Drive authorized';
}

// อัปโหลดรูป (dataURL) ขึ้น Drive → คืน fileId / ถ้าเป็น id เดิมอยู่แล้วก็คืนเดิม
function saveImgToDrive(val) {
  if (!val || String(val).indexOf('data:') !== 0) return val || '';
  const m = String(val).match(/^data:([^;]+);base64,(.*)$/);
  if (!m) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'bd_' + Date.now() + '_' + Math.floor(Math.random() * 1e5));
  const it = DriveApp.getFoldersByName('BreakdownReport_Images');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('BreakdownReport_Images');
  return folder.createFile(blob).getId();
}

// ดึงรูปจาก Drive กลับมาเป็น dataURL (สำหรับแสดง/ใส่ PDF)
function doGetImage(id) {
  try {
    const blob = DriveApp.getFileById(id).getBlob();
    return jsonOut({ success: true, dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) });
  } catch (e) { return jsonOut({ success: false, error: String(e) }); }
}

// เก็บ log ทุกการกระทำลงชีต "_Log" (ใคร/เมื่อไหร่/ทำอะไร)
function writeLog(ss, tracking, action, byName, status) {
  let log = ss.getSheetByName('_Log');
  if (!log) {
    log = ss.insertSheet('_Log');
    const h = log.getRange(1, 1, 1, 5);
    h.setValues([['เวลา', 'เลข Tracking', 'การกระทำ', 'ผู้ดำเนินการ', 'สถานะ']]);
    h.setBackground('#18181b').setFontColor('#ffffff').setFontWeight('bold');
    log.setFrozenRows(1);
  }
  log.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss'),
    tracking || '', action || '', byName || '', status || '',
  ]);
}

// โรงงาน → รหัส CHB (โรงงาน 1 → CHB1, โรงงาน 2 → CHB2)
function factoryToCHB(factory) {
  const m = String(factory || '').match(/(\d+)/);
  return 'CHB' + (m ? m[1] : 'X');
}

function buildPartsStr(parts) {
  return (parts || [])
    .filter(p => p.name)
    .map(p => `${p.name}${p.partNo ? ' (' + p.partNo + ')' : ''} x${p.qty || 1} ${p.unit || 'ชิ้น'}`)
    .join(' | ');
}

// ============================================================
// GET — ดึงข้อมูล (getData / getAll)
// ============================================================
function doGet(e) {
  try {
    const action  = e.parameter.action  || '';
    const year    = e.parameter.year    || String(new Date().getFullYear());
    const factory = (e.parameter.factory || '').replace(/\s+/g, '');
    const area    = e.parameter.area    || '';
    const status  = e.parameter.status  || '';
    const month   = e.parameter.month   || ''; // YYYY-MM
    const machineId = e.parameter.machineId || '';

    if (action === 'login') {
      const role = ROLE_PW[(e.parameter.pw || '').trim()] || '';
      return jsonOut(role ? { success: true, role } : { success: false, error: 'รหัสผ่านไม่ถูกต้อง' });
    }
    if (action === 'getLog') {
      return doGetLog(e.parameter.tracking || '');
    }
    if (action === 'getMachines') {
      return doGetMachines();
    }
    if (action === 'getImage') {
      return doGetImage(e.parameter.id || '');
    }
    if (action === 'getData') {
      return doGetSummary(year, factory, area);
    }
    if (action === 'getAll') {
      return doGetAll(factory, area, status, month, machineId);
    }
    return jsonOut({ success: false, error: 'Unknown action' });

  } catch (err) {
    return jsonOut({ success: false, error: err.toString() });
  }
}

// สรุปข้อมูลรายเดือน (สำหรับ Summary tab)
function doGetSummary(year, factory, area) {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const rows   = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.charAt(0) === '_') return;   // ข้ามชีตภายใน (_Log)
    if (!name.includes('_')) return;
    const sheetMonth = name.split('_').pop();
    if (!sheetMonth.startsWith(year)) return;
    if (factory && !name.startsWith(factory)) return;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue;
      if (area && r[3] !== area) continue;
      rows.push({
        month:       sheetMonth,
        machineName: r[1],  factory:    r[2],
        area:        r[3],  status:     r[6],
        downtimeMin: Number(r[9]) || 0,
        bdType:      r[10],
      });
    }
  });
  return jsonOut({ success: true, year, data: rows });
}

// ดึงข้อมูลทั้งหมด พร้อม rowIndex (สำหรับ Records tab + Edit)
function doGetAll(factory, area, status, month, machineId) {
  const mid = (machineId || '').toLowerCase();
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const rows   = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.charAt(0) === '_') return;   // ข้ามชีตภายใน (_Log)
    if (!name.includes('_')) return;
    const sheetMonth = name.split('_').pop(); // YYYY-MM
    const sheetFactory = name.replace('_' + sheetMonth, '');

    if (factory && sheetFactory !== factory.replace(/\s+/g,'')) return;
    if (month   && sheetMonth !== month) return;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue;
      if (area   && r[3] !== area)   continue;
      if (status && r[6] !== status) continue;
      if (mid    && String(r[4]).toLowerCase().indexOf(mid) < 0) continue; // รหัสเครื่องจักร (บางส่วน)

      const whys = [r[13]||'', r[14]||'', r[15]||'', r[16]||'', r[17]||''].filter(Boolean);
      rows.push({
        rowIndex:    i + 1,            // 1-based row number ใน sheet
        sheetName:   name,
        timestamp:   r[0]  ? String(r[0])  : '',
        machineName: r[1]  || '',
        factory:     r[2]  || '',
        area:        r[3]  || '',
        machineId:   r[4]  || '',
        line:        r[5]  || '',
        status:      r[6]  || '',
        bdStart:     r[7]  || '',
        bdEnd:       r[8]  || '',
        downtimeMin: Number(r[9])  || 0,
        bdType:      r[10] || '',
        problem:     r[11] || '',
        device:      r[12] || '',
        whys,
        corrective:  r[18] || '',
        preventive:  r[19] || '',
        stdLife:     r[20] || '',
        actualLife:  r[21] || '',
        lifeNote:    r[22] || '',
        parts:       r[23] || '',
        tracking:    r[24] || '',
        byName:      r[25] || '',
        acceptedBy:  r[26] || '',
        closedBy:    r[27] || '',
        imgBefore:   r[28] || '',
        imgAfter:    r[29] || '',
      });
    }
  });

  // เรียงล่าสุดก่อน
  rows.sort((a, b) => b.rowIndex - a.rowIndex || b.sheetName.localeCompare(a.sheetName));
  return jsonOut({ success: true, data: rows });
}

// ดึงประวัติ log ของเลข Tracking หนึ่งๆ (สำหรับหน้าดู Log)
function doGetLog(tracking) {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const log = ss.getSheetByName('_Log');
  if (!log) return jsonOut({ success: true, data: [] });
  const data = log.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (tracking && String(r[1]) !== tracking) continue;
    rows.push({ time: String(r[0]), tracking: r[1], action: r[2], byName: r[3], status: r[4] });
  }
  rows.reverse(); // ล่าสุดก่อน
  return jsonOut({ success: true, data: rows });
}

// master รายการเครื่องจักร — อ่านจากแท็บ "_Machines"
// คอลัมน์: A=รหัสเครื่องจักร  B=ชื่อเครื่องจักร  C=โรงงาน  D=พื้นที่  E=ไลน์ (E ไม่บังคับ)
function doGetMachines() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('_Machines');
  if (!sh) return jsonOut({ success: true, data: [] });
  const data = sh.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    rows.push({ id: String(r[0]).trim(), name: r[1] || '', factory: r[2] || '', area: r[3] || '', line: r[4] || '' });
  }
  return jsonOut({ success: true, data: rows });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
