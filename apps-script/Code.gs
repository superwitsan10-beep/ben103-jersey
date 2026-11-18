/**
 * BEN103 Jersey — ระบบพรีออเดอร์ (Google Apps Script + Google Sheets)
 *
 * วิธีติดตั้งอยู่ในไฟล์ README.md
 * - เปิดเป็นเว็บแอป: Deploy > New deployment > Web app
 * - ออเดอร์ลงชีต "Orders", เสื้อแต่ละตัวลงชีต "Items", สรุปยอดในชีต "Summary"
 * - สลิปเก็บในโฟลเดอร์ Google Drive "BEN103 Slips"
 * - เปลี่ยนสถานะออเดอร์ที่คอลัมน์ "สถานะ" ในชีต Orders แล้วลูกค้าจะเห็นตอนเช็คออเดอร์
 */

const CONFIG = {
  SHEET_ID: '',                              // ว่างไว้ถ้าสร้างสคริปต์จากใน Google Sheet (ส่วนขยาย > Apps Script)
  PRICE: 399,                                // ราคาต่อตัว
  SHIP: 50,                                  // ค่าส่งต่อออเดอร์
  NAME_MAX: 20,                              // ความยาวชื่อสกรีนสูงสุด
  MAX_ITEMS: 50,                             // จำนวนเสื้อสูงสุดต่อออเดอร์
  ORDER_PREFIX: 'BEN103-',
  SLIP_FOLDER: 'BEN103 Slips',
  CLOSE_AT: '2026-11-30T23:59:59+07:00',     // ปิดรับอัตโนมัติ (ว่างไว้ = ไม่ปิด)
};

const SIZES = ['SS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', '5XL', '6XL', '7XL', '8XL'];
const STATUSES = ['รอตรวจสลิป', 'ยืนยันแล้ว', 'กำลังผลิต', 'พร้อมรับ/จัดส่ง', 'ยกเลิก'];
const ORDER_HEADERS = ['เวลาสั่ง', 'เลขออเดอร์', 'สถานะ', 'ชื่อ-นามสกุล', 'เบอร์โทร', 'การรับ', 'ที่อยู่', 'อำเภอ/เขต', 'จังหวัด', 'รหัสไปรษณีย์', 'จำนวน (ตัว)', 'รายการเสื้อ', 'ค่าเสื้อ', 'ค่าส่ง', 'ยอดรวม', 'สลิป', 'หมายเหตุ'];
const ITEM_HEADERS = ['เลขออเดอร์', 'ตัวที่', 'ไซซ์', 'ชื่อสกรีน', 'เบอร์', 'สถานะออเดอร์'];
const COL = { phone: 5, zip: 10, status: 3 };   // ตำแหน่งคอลัมน์ในชีต Orders (เริ่มที่ 1)

/* ====================== เว็บแอป ====================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('BEN103 Jersey — พรีออเดอร์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// สำหรับหน้าเว็บที่วางไว้ที่อื่น (เช่น GitHub Pages) เรียกผ่าน fetch
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    let data;
    if (body.action === 'submitOrder') data = submitOrder(body.payload);
    else if (body.action === 'trackOrders') data = trackOrders(body.payload);
    else throw new Error('ไม่รู้จักคำสั่งนี้');
    return json_({ ok: true, data: data });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ====================== ติดตั้งชีต (รันครั้งแรกครั้งเดียว) ====================== */

function setup() {
  const ss = ss_();
  const orders = sheet_('Orders', ORDER_HEADERS);
  const items = sheet_('Items', ITEM_HEADERS);

  orders.getRange(1, COL.phone, orders.getMaxRows(), 1).setNumberFormat('@');
  orders.getRange(1, COL.zip, orders.getMaxRows(), 1).setNumberFormat('@');
  items.getRange(1, 5, items.getMaxRows(), 1).setNumberFormat('@');
  orders.getRange(2, 1, orders.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  orders.getRange(2, 13, orders.getMaxRows() - 1, 3).setNumberFormat('#,##0');

  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
  orders.getRange(2, COL.status, orders.getMaxRows() - 1, 1).setDataValidation(rule);
  orders.setColumnWidth(12, 260);
  orders.getRange(2, 12, orders.getMaxRows() - 1, 1).setWrap(true);

  let sum = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
  sum.clear();
  const rows = [['ไซซ์', 'จำนวนตัว (ไม่รวมยกเลิก)']];
  SIZES.forEach(function (s) { rows.push([s, '=COUNTIFS(Items!C:C,"' + s + '",Items!F:F,"<>ยกเลิก")']); });
  rows.push(['รวมทั้งหมด', '=SUM(B2:B' + (SIZES.length + 1) + ')']);
  rows.push(['', '']);
  rows.push(['จำนวนออเดอร์', '=COUNTIFS(Orders!B2:B,"<>",Orders!C2:C,"<>ยกเลิก")']);
  rows.push(['ยอดเงินรวม (บาท)', '=SUMIFS(Orders!O2:O,Orders!C2:C,"<>ยกเลิก")']);
  rows.push(['ยอดที่ยืนยันแล้ว (บาท)', '=SUMIFS(Orders!O2:O,Orders!C2:C,"<>ยกเลิก",Orders!C2:C,"<>รอตรวจสลิป")']);
  rows.push(['ออเดอร์รอตรวจสลิป', '=COUNTIF(Orders!C2:C,"รอตรวจสลิป")']);
  sum.getRange(1, 1, rows.length, 2).setValues(rows);
  sum.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#135e43').setFontColor('#ffeeaa');
  sum.getRange(SIZES.length + 2, 1, 1, 2).setFontWeight('bold');
  sum.setColumnWidth(1, 200); sum.setColumnWidth(2, 200);

  slipFolder_();
  safeAlert_('ติดตั้งเรียบร้อย ✓ ต่อไป Deploy เป็น Web app ได้เลย');
}

function safeAlert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/* ====================== รับออเดอร์ ====================== */

function submitOrder(p) {
  if (CONFIG.CLOSE_AT && new Date() > new Date(CONFIG.CLOSE_AT)) throw new Error('ขออภัย ปิดรับพรีออเดอร์แล้ว');
  const d = validate_(p);

  // บันทึกสลิปก่อน (นอก lock) แล้วค่อยเปลี่ยนชื่อไฟล์ตามเลขออเดอร์
  const slip = saveSlip_(d.slip, 'pending_' + d.phone + '_' + Date.now());

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const orders = sheet_('Orders', ORDER_HEADERS);
    const items = sheet_('Items', ITEM_HEADERS);

    const n = Math.max(0, orders.getLastRow() - 1) + 1;
    const id = CONFIG.ORDER_PREFIX + ('0000' + n).slice(-4);
    try { slip.setName(id + '_' + d.phone + slip.getName().replace(/^.*(\.\w+)$/, '$1')); } catch (e) {}

    const qty = d.items.length;
    const sub = qty * CONFIG.PRICE;
    const ship = d.del === 'ship' ? CONFIG.SHIP : 0;
    const total = sub + ship;
    const list = d.items.map(function (x, i) { return '#' + (i + 1) + ' ' + x.size + ' · ' + (x.name || '-') + ' · ' + (x.num || '-'); }).join('\n');

    const row = orders.getLastRow() + 1;
    orders.getRange(row, COL.phone).setNumberFormat('@');
    orders.getRange(row, COL.zip).setNumberFormat('@');
    orders.getRange(row, 1, 1, ORDER_HEADERS.length).setValues([[
      new Date(), id, STATUSES[0], safe_(d.full), d.phone,
      d.del === 'ship' ? 'จัดส่ง' : 'รับที่งาน',
      safe_(d.addr.line), safe_(d.addr.district), safe_(d.addr.province), d.addr.zip,
      qty, list, sub, ship, total, slip.getUrl(), ''
    ]]);

    const start = items.getLastRow() + 1;
    items.getRange(start, 5, qty, 1).setNumberFormat('@');
    items.getRange(start, 1, qty, 6).setValues(d.items.map(function (x, i) {
      return [id, i + 1, x.size, x.name, x.num,
        '=IFERROR(INDEX(Orders!C:C,MATCH(A' + (start + i) + ',Orders!B:B,0)),"")'];
    }));
    SpreadsheetApp.flush();
    return { id: id, total: total, qty: qty };
  } finally {
    lock.releaseLock();
  }
}

function validate_(p) {
  if (!p || typeof p !== 'object') throw new Error('ข้อมูลไม่ถูกต้อง');
  const s = function (v) { return String(v == null ? '' : v).trim(); };

  const full = s(p.full).slice(0, 100);
  if (full.length < 2) throw new Error('กรุณากรอกชื่อ-นามสกุล');
  const phone = s(p.phone).replace(/\D/g, '');
  if (!/^0\d{8,9}$/.test(phone)) throw new Error('เบอร์โทรไม่ถูกต้อง');

  const raw = Array.isArray(p.items) ? p.items : [];
  if (!raw.length) throw new Error('กรุณาเลือกเสื้ออย่างน้อย 1 ตัว');
  if (raw.length > CONFIG.MAX_ITEMS) throw new Error('สั่งได้สูงสุด ' + CONFIG.MAX_ITEMS + ' ตัวต่อออเดอร์');
  const items = raw.map(function (x) {
    const size = s(x && x.size).toUpperCase();
    if (SIZES.indexOf(size) < 0) throw new Error('ไซซ์ไม่ถูกต้อง');
    return {
      size: size,
      name: s(x.name).toUpperCase().replace(/[^A-Z .]/g, '').slice(0, CONFIG.NAME_MAX),
      num: s(x.num).replace(/\D/g, '').slice(0, 3),
    };
  });

  const del = p.del === 'ship' ? 'ship' : 'event';
  const addr = { line: '', district: '', province: '', zip: '' };
  if (del === 'ship') {
    const a = p.addr || {};
    addr.line = s(a.line).slice(0, 300);
    addr.district = s(a.district).slice(0, 100);
    addr.province = s(a.province).slice(0, 100);
    addr.zip = s(a.zip).replace(/\D/g, '');
    if (addr.line.length < 5) throw new Error('กรุณากรอกที่อยู่จัดส่ง');
    if (!addr.province) throw new Error('กรุณาเลือกจังหวัด');
    if (addr.district.length < 2) throw new Error('กรุณากรอกอำเภอ/เขต');
    if (!/^\d{5}$/.test(addr.zip)) throw new Error('รหัสไปรษณีย์ต้องมี 5 หลัก');
  }

  if (!p.slip) throw new Error('กรุณาแนบสลิปการโอน');
  return { full: full, phone: phone, items: items, del: del, addr: addr, slip: String(p.slip) };
}

// กันข้อความที่ขึ้นต้นด้วย = + - @ ไม่ให้กลายเป็นสูตรในชีต
function safe_(v) {
  v = String(v == null ? '' : v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function saveSlip_(dataUrl, name) {
  const m = /^data:(image\/[\w.+-]+);base64,(.+)$/.exec(dataUrl || '');
  if (!m) throw new Error('ไฟล์สลิปต้องเป็นรูปภาพ');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 15 * 1024 * 1024) throw new Error('ไฟล์สลิปใหญ่เกิน 15MB');
  const ext = (m[1].split('/')[1] || 'jpg').replace('jpeg', 'jpg');
  return slipFolder_().createFile(Utilities.newBlob(bytes, m[1], name + '.' + ext));
}

function slipFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SLIP_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(CONFIG.SLIP_FOLDER);
  props.setProperty('SLIP_FOLDER_ID', f.getId());
  return f;
}

/* ====================== เช็คออเดอร์ ====================== */

function trackOrders(p) {
  const key = String((p && p.phone) || '').replace(/\D/g, '').replace(/^0+/, '');
  if (key.length < 8) throw new Error('เบอร์โทรไม่ถูกต้อง');

  const orders = sheet_('Orders', ORDER_HEADERS);
  const items = sheet_('Items', ITEM_HEADERS);
  const ov = orders.getLastRow() > 1 ? orders.getRange(2, 1, orders.getLastRow() - 1, ORDER_HEADERS.length).getValues() : [];
  const found = ov.filter(function (r) { return String(r[COL.phone - 1]).replace(/\D/g, '').replace(/^0+/, '') === key; });
  if (!found.length) return [];

  const ids = {};
  found.forEach(function (r) { ids[r[1]] = []; });
  const iv = items.getLastRow() > 1 ? items.getRange(2, 1, items.getLastRow() - 1, 5).getValues() : [];
  iv.forEach(function (r) { if (ids[r[0]]) ids[r[0]].push({ size: String(r[2]), name: String(r[3]), num: String(r[4]) }); });

  return found.reverse().map(function (r) {
    const status = String(r[COL.status - 1]);
    return { id: String(r[1]), status: status, step: Math.max(0, STATUSES.indexOf(status)), total: Number(r[14]) || 0, items: ids[r[1]] };
  });
}

/* ====================== helpers ====================== */

function ss_() {
  return CONFIG.SHEET_ID ? SpreadsheetApp.openById(CONFIG.SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#135e43').setFontColor('#ffeeaa');
    sh.setFrozenRows(1);
  }
  return sh;
}
