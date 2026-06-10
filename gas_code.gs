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

const SPREADSHEET_ID = ''; // ← ใส่ ID ของ Google Sheet ที่ต้องการ

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
];

// ============================================================
// POST — บันทึก / อัปเดต
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ---- UPDATE existing row ----
    if (data.action === 'update') {
      const sheet = ss.getSheetByName(data.sheetName);
      if (!sheet || !data.rowIndex) throw new Error('Sheet or rowIndex not found');

      const whys    = data.whys || [];
      const partsStr = buildPartsStr(data.parts);
      const row = buildRow(data, whys, partsStr, /*keepTimestamp=*/true);
      sheet.getRange(data.rowIndex, 1, 1, row.length).setValues([row]);

      return jsonOut({ success: true, action: 'updated' });
    }

    // ---- CREATE new row ----
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

    const whys    = data.whys || [];
    const partsStr = buildPartsStr(data.parts);
    sheet.appendRow(buildRow(data, whys, partsStr, /*keepTimestamp=*/false, now));

    return jsonOut({ success: true, sheet: sheetName });

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
  ];
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

    if (action === 'getData') {
      return doGetSummary(year, factory, area);
    }
    if (action === 'getAll') {
      return doGetAll(factory, area, status, month);
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
function doGetAll(factory, area, status, month) {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const rows   = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
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
      });
    }
  });

  // เรียงล่าสุดก่อน
  rows.sort((a, b) => b.rowIndex - a.rowIndex || b.sheetName.localeCompare(a.sheetName));
  return jsonOut({ success: true, data: rows });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
