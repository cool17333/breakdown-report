// ===================================================
// Breakdown Report — SharePoint Server
// รันด้วย: node server.js
// ===================================================
require('dotenv').config();
require('isomorphic-fetch');

const express = require('express');
const cors    = require('cors');
const XLSX    = require('xlsx');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client, ResponseType } = require('@microsoft/microsoft-graph-client');

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(cors({ origin: '*' }));
app.use(express.json({ limit: '20mb' }));

// ============================================================
// AZURE AD — CLIENT CREDENTIALS (App-only, ไม่ต้อง login)
// ============================================================
const msalApp = new ConfidentialClientApplication({
    auth: {
        clientId:     process.env.CLIENT_ID,
        authority:    `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET,
    },
});

async function getGraphClient() {
    const { accessToken } = await msalApp.acquireTokenByClientCredential({
        scopes: ['https://graph.microsoft.com/.default'],
    });
    return Client.init({
        authProvider: (done) => done(null, accessToken),
    });
}

// ============================================================
// CONFIG
// ============================================================
const SITE_ID     = process.env.SHAREPOINT_SITE_ID;
const DRIVE_ID    = process.env.SHAREPOINT_DRIVE_ID;
const BASE_FOLDER = process.env.BASE_FOLDER || 'Breakdown Report';

const HEADERS = [
    'วันที่บันทึก', 'ชื่อเครื่องจักร', 'โรงงาน', 'พื้นที่',
    'รหัสเครื่องจักร', 'สาย / ตำแหน่ง', 'สถานะ',
    'เวลาเริ่ม Breakdown', 'เวลาเสร็จสิ้น', 'Downtime (นาที)',
    'ประเภท Breakdown',
    'ปัญหาที่พบ', 'อุปกรณ์ที่เกิดปัญหา',
    'Why 1', 'Why 2', 'Why 3', 'Why 4', 'Why 5',
    'มาตรการแก้ไข', 'มาตรการป้องกัน',
    'อายุมาตรฐาน (Std.)', 'อายุจริงตอนเสีย', 'หมายเหตุอายุ',
    'อะไหล่ที่ใช้',
];

// ============================================================
// GRAPH API HELPERS
// ============================================================
function driveItemPath(client, itemPath) {
    return client.api(`/sites/${SITE_ID}/drives/${DRIVE_ID}/root:/${itemPath}`);
}

// ดาวน์โหลดไฟล์เป็น Buffer
async function downloadBuffer(client, itemPath) {
    const raw = await client
        .api(`/sites/${SITE_ID}/drives/${DRIVE_ID}/root:/${itemPath}:/content`)
        .responseType(ResponseType.BUFFER)
        .get();
    return Buffer.from(raw);
}

// อัปโหลด Buffer ขึ้น SharePoint (สร้าง/ทับไฟล์เดิม)
async function uploadBuffer(client, itemPath, buffer) {
    await client
        .api(`/sites/${SITE_ID}/drives/${DRIVE_ID}/root:/${itemPath}:/content`)
        .put(buffer);
}

// สร้าง folder (ถ้ายังไม่มี)
async function ensureFolder(client, folderPath) {
    try {
        await driveItemPath(client, folderPath).get();
    } catch {
        const parts    = folderPath.split('/');
        const name     = parts.pop();
        const parentPath = parts.join('/');
        const endpoint = parentPath
            ? `/sites/${SITE_ID}/drives/${DRIVE_ID}/root:/${parentPath}:/children`
            : `/sites/${SITE_ID}/drives/${DRIVE_ID}/root/children`;
        await client.api(endpoint).post({
            name,
            folder: {},
            '@microsoft.graph.conflictBehavior': 'ignore',
        });
    }
}

// สร้าง Excel ใหม่พร้อม header
function createNewWorkbook() {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([HEADERS]);

    // ความกว้าง column
    ws['!cols'] = HEADERS.map((_, i) => ({
        wch: [15, 20, 12, 16, 16, 16, 18, 19, 19, 14, 18, 14, 14, 14, 14, 30, 30, 25, 25, 25, 25, 25, 30, 30, 16, 16, 20, 35][i] || 20,
    }));

    // Freeze header row
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };

    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    return wb;
}

// ============================================================
// GET or CREATE Excel ไฟล์ใน SharePoint
// โครงสร้าง: Breakdown Report / {factoryCode} / {YYYY-MM}.xlsx
// ============================================================
async function getOrCreateExcel(client, factoryCode, yearMonth) {
    const folderPath = `${BASE_FOLDER}/${factoryCode}`;
    const filePath   = `${folderPath}/${yearMonth}.xlsx`;

    try {
        // ไฟล์มีอยู่แล้ว
        const buffer = await downloadBuffer(client, filePath);
        console.log(`📂 Found existing file: ${filePath}`);
        return { buffer, filePath, isNew: false };
    } catch (err) {
        if (err.statusCode !== 404) throw err;

        // สร้าง folder + ไฟล์ใหม่
        console.log(`📁 Creating folder: ${folderPath}`);
        await ensureFolder(client, BASE_FOLDER);
        await ensureFolder(client, folderPath);

        const wb     = createNewWorkbook();
        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        console.log(`📄 Creating new file: ${filePath}`);
        await uploadBuffer(client, filePath, buffer);

        return { buffer, filePath, isNew: true };
    }
}

// ============================================================
// POST /add-data — บันทึก Breakdown ลง Excel
// ============================================================
app.post('/add-data', async (req, res) => {
    try {
        const data        = req.body;
        const factoryCode = (data.factory || 'Unknown').replace(/\s+/g, '');
        const yearMonth   = new Date().toISOString().slice(0, 7); // YYYY-MM
        const filePath    = `${BASE_FOLDER}/${factoryCode}/${yearMonth}.xlsx`;

        console.log(`\n📥 Add Data: ${data.machineName} | ${data.factory} | ${yearMonth}`);

        const client = await getGraphClient();
        const { buffer } = await getOrCreateExcel(client, factoryCode, yearMonth);

        // Parse → append row → serialize
        const wb  = XLSX.read(buffer, { type: 'buffer', cellDates: true });
        const ws  = wb.Sheets['Data'];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1 });

        const whys    = data.whys || [];
        const partsStr = (data.parts || [])
            .filter(p => p.name)
            .map(p => `${p.name}${p.partNo ? ` (${p.partNo})` : ''} x${p.qty || 1} ${p.unit || 'ชิ้น'}`)
            .join(' | ');

        const newRow = [
            new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' }),
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
            whys[0] || '', whys[1] || '', whys[2] || '', whys[3] || '', whys[4] || '',
            data.corrective    || '',
            data.preventive    || '',
            data.stdLife       || '',
            data.actualLife    || '',
            data.lifeNote      || '',
            partsStr,
        ];

        aoa.push(newRow);

        const newWs     = XLSX.utils.aoa_to_sheet(aoa);
        newWs['!cols']  = ws['!cols'];
        newWs['!freeze'] = { xSplit: 0, ySplit: 1 };
        wb.Sheets['Data'] = newWs;

        const newBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        await uploadBuffer(client, filePath, newBuffer);

        const totalRows = aoa.length - 1; // ไม่นับ header
        console.log(`✅ Saved → ${filePath} (แถวที่ ${totalRows})`);

        res.json({
            success:  true,
            file:     filePath,
            row:      totalRows,
            message:  `บันทึกสำเร็จ (แถวที่ ${totalRows} ใน ${yearMonth}.xlsx)`,
        });

    } catch (err) {
        console.error('❌ /add-data error:', err.message || err);
        res.status(500).json({ success: false, error: err.message || 'Unknown error' });
    }
});

// ============================================================
// GET /summary?year=2026&factory=โรงงาน1&area=Frozenดิบ
// ============================================================
app.get('/summary', async (req, res) => {
    try {
        const year        = req.query.year    || String(new Date().getFullYear());
        const factory     = req.query.factory || '';
        const area        = req.query.area    || '';
        const factoryCode = factory.replace(/\s+/g, '');

        console.log(`\n📊 Summary: year=${year} factory=${factoryCode} area=${area}`);

        const client = await getGraphClient();
        const allFiles = [];

        if (factoryCode) {
            // โหลดจาก folder โรงงานเดียว
            try {
                const folderPath = `${BASE_FOLDER}/${factoryCode}`;
                const items = await client
                    .api(`/sites/${SITE_ID}/drives/${DRIVE_ID}/root:/${folderPath}:/children`)
                    .get();
                allFiles.push(...items.value
                    .filter(i => i.name.endsWith('.xlsx') && i.name.includes(year))
                    .map(i => ({ ...i, factoryCode }))
                );
            } catch { /* folder doesn't exist yet */ }
        } else {
            // โหลดจากทุก folder ใน BASE_FOLDER
            try {
                const folders = await client
                    .api(`/sites/${SITE_ID}/drives/${DRIVE_ID}/root:/${BASE_FOLDER}:/children`)
                    .get();
                for (const folder of folders.value.filter(i => i.folder)) {
                    try {
                        const items = await client
                            .api(`/sites/${SITE_ID}/drives/${DRIVE_ID}/items/${folder.id}/children`)
                            .get();
                        allFiles.push(...items.value
                            .filter(i => i.name.endsWith('.xlsx') && i.name.includes(year))
                            .map(i => ({ ...i, factoryCode: folder.name }))
                        );
                    } catch { /* skip empty folders */ }
                }
            } catch { /* BASE_FOLDER doesn't exist */ }
        }

        const allRows = [];

        for (const file of allFiles) {
            const month = file.name.replace('.xlsx', ''); // YYYY-MM
            const raw = await client
                .api(`/sites/${SITE_ID}/drives/${DRIVE_ID}/items/${file.id}/content`)
                .responseType(ResponseType.BUFFER)
                .get();
            const buf  = Buffer.from(raw);
            const wb   = XLSX.read(buf, { type: 'buffer' });
            const rows = XLSX.utils.sheet_to_json(wb.Sheets['Data'], { header: 1 });

            for (let i = 1; i < rows.length; i++) {
                const r = rows[i];
                if (!r || !r[0]) continue;
                if (area && r[3] !== area) continue;

                allRows.push({
                    month,
                    machineName: r[1]  || '',
                    factory:     r[2]  || '',
                    area:        r[3]  || '',
                    status:      r[6]  || '',
                    downtimeMin: Number(r[9])  || 0,
                    bdType:      r[10] || '',
                });
            }
        }

        console.log(`✅ Summary: ${allRows.length} records from ${allFiles.length} files`);
        res.json({ success: true, year, data: allRows });

    } catch (err) {
        console.error('❌ /summary error:', err.message || err);
        res.status(500).json({ success: false, error: err.message || 'Unknown error' });
    }
});

// ============================================================
// GET /health — ตรวจสอบ server + SharePoint connection
// ============================================================
app.get('/health', async (req, res) => {
    try {
        const client = await getGraphClient();
        const site   = await client.api(`/sites/${SITE_ID}`).select('displayName,webUrl').get();
        res.json({
            status:    'ok',
            server:    `localhost:${PORT}`,
            sharePoint: site.displayName,
            siteUrl:   site.webUrl,
            baseFolder: BASE_FOLDER,
        });
    } catch (err) {
        res.status(500).json({ status: 'error', error: err.message });
    }
});

// ============================================================
// START SERVER
// ============================================================
app.listen(PORT, () => {
    console.log('╔══════════════════════════════════════════╗');
    console.log(`║  Breakdown Server — http://localhost:${PORT}  ║`);
    console.log('╠══════════════════════════════════════════╣');
    console.log(`║  POST /add-data    บันทึก Breakdown       ║`);
    console.log(`║  GET  /summary     ดึงข้อมูลสรุป          ║`);
    console.log(`║  GET  /health      ตรวจสอบการเชื่อมต่อ   ║`);
    console.log('╚══════════════════════════════════════════╝');
});
