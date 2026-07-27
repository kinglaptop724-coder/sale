/**
 * ==========================================================
 * ตลาดนัดนักเรียน (Student Market Flip) - Backend
 * Google Apps Script + Google Sheet เป็นฐานข้อมูล
 * ==========================================================
 *
 * วิธีติดตั้ง:
 * 1. สร้าง Google Sheet ใหม่ (ไฟล์เปล่าก็ได้)
 * 2. เมนู Extensions > Apps Script
 * 3. ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดไฟล์นี้ทั้งไฟล์
 * 4. เลือกฟังก์ชัน setup แล้วกด Run (ครั้งแรกจะขอ authorize ให้กด Allow)
 *    -> จะสร้างชีตและข้อมูลตั้งต้นให้อัตโนมัติ
 * 5. Deploy > New deployment > เลือกประเภท "Web app"
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. คัดลอก Web app URL ไปใส่ในไฟล์ index.html (ตัวแปร API_URL)
 * 7. กลับมาที่ Google Sheet แท็บ "Config" ตั้งค่ารหัสห้อง (JOIN_CODE) ได้ตามต้องการ
 * 8. เมื่อพร้อมเริ่มแข่ง ให้ไปที่เมนู "เกมตลาดนัด" > "เริ่มเกม" บน Google Sheet
 */

const SHEETS = {
  CONFIG: 'Config',
  ITEMS: 'Items',
  STUDENTS: 'Students',
  INVENTORY: 'Inventory',
  LISTINGS: 'Listings',
  TRANSACTIONS: 'Transactions',
  SHOPS: 'Shops'
};

// ---------- TRUCK STYLES (ธีมรถขายของให้นักเรียนเลือก) ----------
const TRUCK_STYLES = [
  { id: 'red',    label: 'คลาสสิกแดง',     emoji: '🚚', colorA: '#d1453a', colorB: '#f7ded9' },
  { id: 'green',  label: 'สวนผลไม้เขียว',   emoji: '🚐', colorA: '#2f8f6f', colorB: '#dcf3e9' },
  { id: 'pink',   label: 'ขนมหวานชมพู',    emoji: '🍦', colorA: '#c94f8a', colorB: '#f8dfec' },
  { id: 'blue',   label: 'ทะเลฟ้าคราม',    emoji: '🚛', colorA: '#3f6fbf', colorB: '#dde7f7' },
  { id: 'gold',   label: 'หรูหราทอง',      emoji: '🛻', colorA: '#c98a2c', colorB: '#f8ecd6' },
  { id: 'purple', label: 'ราตรีม่วง',      emoji: '🚙', colorA: '#7f5aa2', colorB: '#ecdff5' }
];
function getTruckStyle(styleId) {
  return TRUCK_STYLES.find(t => t.id === styleId) || TRUCK_STYLES[0];
}

// ---------- SETUP ----------
function setupCore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const config = getOrCreateSheet(ss, SHEETS.CONFIG);
  config.clear();
  config.getRange('A1:B10').setValues([
    ['KEY', 'VALUE'],
    ['GAME_NAME', 'ตลาดนัดนักเรียน'],
    ['JOIN_CODE', 'ROOM1'],
    ['START_CASH', 500],
    ['ROUND_MINUTES', 30],
    ['GAME_STARTED_AT', ''],
    ['GAME_ENDED', 'FALSE'],
    ['LISTING_FEE_PERCENT', 0],
    ['TRUCK_COST', 150],
    ['TEACHER_KEY', 'teacher123']
  ]);

  const items = getOrCreateSheet(ss, SHEETS.ITEMS);
  items.clear();
  items.getRange('A1:D1').setValues([['itemName', 'basePrice', 'volatility', 'emoji']]);
  items.getRange('A2:D7').setValues([
    ['มะม่วง', 20, 0.4, '🥭'],
    ['ตุ๊กตา', 80, 0.3, '🧸'],
    ['สมุด', 15, 0.2, '📓'],
    ['เสื้อยืด', 120, 0.35, '👕'],
    ['ขนมถุง', 10, 0.5, '🍪'],
    ['หูฟัง', 250, 0.25, '🎧']
  ]);

  const students = getOrCreateSheet(ss, SHEETS.STUDENTS);
  students.clear();
  students.getRange('A1:E1').setValues([['id', 'name', 'cash', 'joinedAt', 'lastSeen']]);

  const inv = getOrCreateSheet(ss, SHEETS.INVENTORY);
  inv.clear();
  inv.getRange('A1:D1').setValues([['studentId', 'itemName', 'qty', 'avgCost']]);

  const listings = getOrCreateSheet(ss, SHEETS.LISTINGS);
  listings.clear();
  listings.getRange('A1:G1').setValues([['id', 'sellerId', 'sellerName', 'itemName', 'qty', 'price', 'createdAt']]);

  const tx = getOrCreateSheet(ss, SHEETS.TRANSACTIONS);
  tx.clear();
  tx.getRange('A1:H1').setValues([['id', 'time', 'buyerId', 'buyerName', 'sellerId', 'sellerName', 'itemName', 'qtyPriceTotal']]);

  const shops = getOrCreateSheet(ss, SHEETS.SHOPS);
  shops.clear();
  shops.getRange('A1:F1').setValues([['studentId', 'shopName', 'styleId', 'isOpen', 'openedAt', 'tagline']]);
}

function setup() {
  setupCore();
  SpreadsheetApp.getUi().alert('ติดตั้งเสร็จแล้ว! ไปที่ Deploy > New deployment เพื่อสร้าง Web App URL');
}

function getOrCreateSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// เมนูช่วยครูคุมเกมจากหน้า Google Sheet โดยตรง
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('เกมตลาดนัด')
    .addItem('เริ่มเกม', 'startGame')
    .addItem('จบเกม', 'endGame')
    .addItem('รีเซ็ตเกมทั้งหมด (ล้างข้อมูล)', 'setup')
    .addToUi();
}

function startGameCore() {
  setConfig('GAME_STARTED_AT', new Date().toISOString());
  setConfig('GAME_ENDED', 'FALSE');
}

function endGameCore() {
  setConfig('GAME_ENDED', 'TRUE');
}

function startGame() {
  startGameCore();
  SpreadsheetApp.getUi().alert('เริ่มเกมแล้ว!');
}

function endGame() {
  endGameCore();
  SpreadsheetApp.getUi().alert('จบเกมแล้ว!');
}

// ---------- CONFIG HELPERS ----------
function getConfigMap() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.CONFIG);
  const values = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    map[values[i][0]] = values[i][1];
  }
  return map;
}

function setConfig(key, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.CONFIG);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

// action ที่แค่ "อ่าน" ข้อมูล ไม่ได้แก้ไขอะไร ไม่จำเป็นต้องใช้ Lock เลย
// (สำคัญมาก: หน้าเว็บนักเรียนแต่ละคน poll 'state' ทุก 3 วิ ถ้าบังคับให้ทุกคนต้องแย่ง
//  Lock เดียวกันแม้แต่ตอนแค่อ่านข้อมูล คิวจะยาวมากเมื่อมีคนเล่นพร้อมกันหลายคน
//  จนทำให้ action อื่น เช่น ซื้อ/ขาย/เปิดร้าน ต้องรอนานเกินจน Lock timeout)
const READ_ONLY_ACTIONS = ['state', 'teacherState', 'teacherGetConfig'];

// ---------- ENTRY POINT ----------
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (READ_ONLY_ACTIONS.indexOf(action) !== -1) {
      return jsonResponse(runAction(action, body));
    }

    // action ที่เหลือ = เขียน/แก้ไขข้อมูล ต้องกันชนกันด้วย Lock
    const lock = LockService.getScriptLock();
    let gotLock = false;
    try {
      // สำคัญ: waitLock ต้องอยู่ใน try ด้วย ไม่งั้นถ้าแย่งล็อกไม่ได้
      // มันจะ throw ออกไปนอก try/catch ทำให้ Apps Script ตอบกลับเป็นหน้า error (ไม่ใช่ JSON)
      // ฝั่งหน้าเว็บจะ parse JSON ไม่ได้ แล้วค้างตลอดไป (ปุ่มไม่ถูกปลดล็อก)
      lock.waitLock(20000);
      gotLock = true;
      return jsonResponse(runAction(action, body));
    } finally {
      if (gotLock) lock.releaseLock();
    }
  } catch (err) {
    // ไม่ว่าจะ error จากขั้นตอนไหน (รวมถึงแย่งล็อกไม่ได้) ก็ยังส่ง JSON กลับเสมอ
    // เพื่อให้หน้าเว็บโชว์ error แล้วปลดล็อกปุ่มได้ ไม่ค้าง
    return jsonResponse({ error: 'เซิร์ฟเวอร์ไม่ว่าง กรุณาลองใหม่อีกครั้ง (' + String(err) + ')' });
  }
}

function runAction(action, body) {
  switch (action) {
    case 'join': return handleJoin(body);
    case 'state': return handleState(body);
    case 'teacherState': return handleTeacherState(body);
    case 'teacherStart': return handleTeacherStart(body);
    case 'teacherEnd': return handleTeacherEnd(body);
    case 'teacherReset': return handleTeacherReset(body);
    case 'teacherGetConfig': return handleTeacherGetConfig(body);
    case 'teacherUpdateConfig': return handleTeacherUpdateConfig(body);
    case 'buyFromSupplier': return handleBuyFromSupplier(body);
    case 'openShop': return handleOpenShop(body);
    case 'updateShop': return handleUpdateShop(body);
    case 'listForSale': return handleListForSale(body);
    case 'cancelListing': return handleCancelListing(body);
    case 'buyListing': return handleBuyListing(body);
    default: return { error: 'ไม่รู้จัก action: ' + action };
  }
}

function doGet(e) {
  return jsonResponse({ ok: true, message: 'Student Market Flip API ทำงานอยู่' });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- SHEET UTILS ----------
function sheet(name) {
  return SpreadsheetApp.getActive().getSheetByName(name);
}

function rowsToObjects(sh) {
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const obj = {};
    headers.forEach((h, idx) => obj[h] = values[i][idx]);
    obj._row = i + 1; // 1-indexed sheet row, for updates
    out.push(obj);
  }
  return out;
}

function genId(prefix) {
  return prefix + '_' + Utilities.getUuid().slice(0, 8);
}

// ---------- SUPPLIER PRICE MODEL ----------
// ราคาสินค้าจากส่วนกลางแกว่งขึ้นลงตามเวลาแบบ deterministic (ไม่ต้องเก็บ state)
// ใช้ sine wave ผสม noise เบาๆ จาก hash ชื่อสินค้า เพื่อให้แต่ละสินค้าแกว่งไม่พร้อมกัน
function computeSupplierPrice(itemName, basePrice, volatility) {
  const now = new Date().getTime();
  const minutesElapsed = now / 60000;
  let seed = 0;
  for (let i = 0; i < itemName.length; i++) seed += itemName.charCodeAt(i);
  const phase = seed % 60;
  const wave = Math.sin((minutesElapsed + phase) / 4);
  const price = basePrice * (1 + volatility * wave * 0.5);
  return Math.max(1, Math.round(price));
}

function getSupplierItems() {
  const items = rowsToObjects(sheet(SHEETS.ITEMS));
  return items.map(it => ({
    itemName: it.itemName,
    price: computeSupplierPrice(it.itemName, it.basePrice, it.volatility),
    emoji: it.emoji
  }));
}

function getSupplierPriceMap() {
  const map = {};
  getSupplierItems().forEach(it => map[it.itemName] = it.price);
  return map;
}

// ---------- STUDENT HELPERS ----------
function findStudent(studentId) {
  const students = rowsToObjects(sheet(SHEETS.STUDENTS));
  return students.find(s => s.id === studentId);
}

function updateStudentCash(studentId, newCash) {
  const sh = sheet(SHEETS.STUDENTS);
  const students = rowsToObjects(sh);
  const s = students.find(st => st.id === studentId);
  if (!s) throw new Error('ไม่พบผู้เล่น');
  sh.getRange(s._row, 3).setValue(newCash); // column C = cash
}

function getInventory(studentId) {
  const inv = rowsToObjects(sheet(SHEETS.INVENTORY));
  return inv.filter(i => i.studentId === studentId && i.qty > 0);
}

function addInventory(studentId, itemName, qty, unitCost) {
  const sh = sheet(SHEETS.INVENTORY);
  const rows = rowsToObjects(sh);
  const existing = rows.find(r => r.studentId === studentId && r.itemName === itemName);
  if (existing) {
    const newQty = existing.qty + qty;
    const newAvgCost = ((existing.qty * existing.avgCost) + (qty * unitCost)) / newQty;
    sh.getRange(existing._row, 3, 1, 2).setValues([[newQty, newAvgCost]]);
  } else {
    sh.appendRow([studentId, itemName, qty, unitCost]);
  }
}

function removeInventory(studentId, itemName, qty) {
  const sh = sheet(SHEETS.INVENTORY);
  const rows = rowsToObjects(sh);
  const existing = rows.find(r => r.studentId === studentId && r.itemName === itemName);
  if (!existing || existing.qty < qty) throw new Error('สินค้าในสต๊อกไม่พอ');
  sh.getRange(existing._row, 3).setValue(existing.qty - qty);
  return existing.avgCost;
}

// ---------- SHOP (รถขายของ) HELPERS ----------
// สร้างชีต Shops อัตโนมัติถ้ายังไม่มี (กรณีอัปเดตโค้ดโดยไม่ได้รัน setup() ใหม่ เพื่อไม่ให้ข้อมูลเกมเดิมหาย)
function ensureShopsSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEETS.SHOPS);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.SHOPS);
    sh.getRange('A1:F1').setValues([['studentId', 'shopName', 'styleId', 'isOpen', 'openedAt', 'tagline']]);
  } else if (sh.getLastColumn() < 6) {
    // ชีตเก่าที่ยังไม่มีคอลัมน์ tagline (สร้างก่อนอัปเดตฟีเจอร์นี้) -> เพิ่มคอลัมน์ให้อัตโนมัติ
    sh.getRange('F1').setValue('tagline');
  }
  return sh;
}

function getTruckCost() {
  const cfg = getConfigMap();
  const n = Number(cfg.TRUCK_COST);
  return (cfg.TRUCK_COST === undefined || cfg.TRUCK_COST === '' || isNaN(n)) ? 150 : n;
}

function getTeacherKey() {
  const cfg = getConfigMap();
  return (cfg.TEACHER_KEY === undefined || cfg.TEACHER_KEY === '') ? 'teacher123' : String(cfg.TEACHER_KEY);
}

function getShop(studentId) {
  const shops = rowsToObjects(ensureShopsSheet());
  return shops.find(s => s.studentId === studentId);
}

// Google Sheets จะแปลงค่า string "TRUE"/"FALSE" ที่เขียนเข้าไปให้กลายเป็นค่า Boolean จริงโดยอัตโนมัติ
// เมื่ออ่านกลับมาผ่าน getValues() จึงอาจได้ boolean true หรือ string 'TRUE' ก็ได้ (ขึ้นกับว่า Sheets ตีความยังไง)
// ฟังก์ชันนี้เช็คให้ครอบคลุมทั้งสองแบบ กันปัญหา "เปิดร้านแล้วระบบไม่รู้ว่าเปิด"
function isOpenValue(v) {
  return v === true || v === 'TRUE' || v === 'true';
}

function isShopOpen(shop) {
  return !!shop && isOpenValue(shop.isOpen);
}

function saveShop(studentId, shopName, styleId, isOpen, tagline) {
  const sh = ensureShopsSheet();
  const rows = rowsToObjects(sh);
  const existing = rows.find(r => r.studentId === studentId);
  const openedAt = (existing && existing.openedAt) ? existing.openedAt : new Date().toISOString();
  const taglineValue = tagline !== undefined ? tagline : (existing ? (existing.tagline || '') : '');
  if (existing) {
    sh.getRange(existing._row, 1, 1, 6).setValues([[studentId, shopName, styleId, isOpen ? 'TRUE' : 'FALSE', openedAt, taglineValue]]);
  } else {
    sh.appendRow([studentId, shopName, styleId, isOpen ? 'TRUE' : 'FALSE', openedAt, taglineValue]);
  }
}

function getAllShopsMap() {
  const shops = rowsToObjects(ensureShopsSheet());
  const map = {};
  shops.forEach(s => { map[s.studentId] = s; });
  return map;
}

// ---------- ACTION HANDLERS ----------
function handleJoin(body) {
  const cfg = getConfigMap();
  if (body.code !== cfg.JOIN_CODE) return { error: 'รหัสห้องไม่ถูกต้อง' };

  const sh = sheet(SHEETS.STUDENTS);
  const students = rowsToObjects(sh);
  let existing = students.find(s => s.name === body.name);
  if (existing) {
    sh.getRange(existing._row, 5).setValue(new Date().toISOString());
    return { studentId: existing.id, name: existing.name, cash: existing.cash };
  }
  const id = genId('stu');
  sh.appendRow([id, body.name, cfg.START_CASH, new Date().toISOString(), new Date().toISOString()]);
  return { studentId: id, name: body.name, cash: cfg.START_CASH };
}

function handleState(body) {
  const cfg = getConfigMap();
  const student = findStudent(body.studentId);
  if (!student) return { error: 'ไม่พบผู้เล่น กรุณาเข้าเกมใหม่' };

  const priceMap = getSupplierPriceMap();
  const inventory = getInventory(body.studentId);
  const allListings = rowsToObjects(sheet(SHEETS.LISTINGS)).filter(l => l.qty > 0);
  const shopsMap = getAllShopsMap();
  const myShop = shopsMap[body.studentId];

  // มูลค่าสุทธิ = เงินสด + มูลค่าสต๊อกตามราคาตลาดปัจจุบัน
  const students = rowsToObjects(sheet(SHEETS.STUDENTS));
  const invAll = rowsToObjects(sheet(SHEETS.INVENTORY));
  const listingsAll = allListings;

  const leaderboard = students.map(s => {
    const myInv = invAll.filter(i => i.studentId === s.id);
    const invValue = myInv.reduce((sum, i) => sum + (i.qty * (priceMap[i.itemName] || 0)), 0);
    const myListingsValue = listingsAll.filter(l => l.sellerId === s.id)
      .reduce((sum, l) => sum + (l.qty * l.price), 0);
    return {
      name: s.name,
      cash: s.cash,
      netWorth: Math.round(s.cash + invValue + myListingsValue)
    };
  }).sort((a, b) => b.netWorth - a.netWorth);

  let timeLeftSeconds = null;
  if (cfg.GAME_STARTED_AT) {
    const started = new Date(cfg.GAME_STARTED_AT).getTime();
    const elapsedSec = (Date.now() - started) / 1000;
    timeLeftSeconds = Math.max(0, Math.round(cfg.ROUND_MINUTES * 60 - elapsedSec));
  }

  return {
    gameName: cfg.GAME_NAME,
    gameEnded: isOpenValue(cfg.GAME_ENDED) || (timeLeftSeconds !== null && timeLeftSeconds <= 0),
    timeLeftSeconds,
    student: { id: student.id, name: student.name, cash: student.cash },
    supplierItems: getSupplierItems(),
    inventory: inventory.map(i => ({ itemName: i.itemName, qty: i.qty, avgCost: Math.round(i.avgCost * 100) / 100 })),
    myListings: allListings.filter(l => l.sellerId === body.studentId)
      .map(l => ({ id: l.id, itemName: l.itemName, qty: l.qty, price: l.price })),
    market: allListings.filter(l => l.sellerId !== body.studentId)
      .map(l => {
        const sellerShop = shopsMap[l.sellerId];
        return {
          id: l.id,
          sellerId: l.sellerId,
          sellerName: l.sellerName,
          shopName: (sellerShop && isOpenValue(sellerShop.isOpen)) ? sellerShop.shopName : ('แผง' + l.sellerName),
          styleId: (sellerShop && isOpenValue(sellerShop.isOpen)) ? sellerShop.styleId : null,
          shopTagline: (sellerShop && isOpenValue(sellerShop.isOpen)) ? (sellerShop.tagline || '') : '',
          itemName: l.itemName, qty: l.qty, price: l.price
        };
      }),
    leaderboard: leaderboard.slice(0, 15),
    truckCost: getTruckCost(),
    truckStyles: TRUCK_STYLES,
    myShop: {
      opened: isShopOpen(myShop),
      name: myShop ? myShop.shopName : '',
      styleId: myShop ? myShop.styleId : TRUCK_STYLES[0].id,
      tagline: myShop ? (myShop.tagline || '') : ''
    }
  };
}

function handleTeacherState(body) {
  const cfg = getConfigMap();
  if (String(body.key || '') !== getTeacherKey()) return { error: 'รหัสครูไม่ถูกต้อง' };

  const students = rowsToObjects(sheet(SHEETS.STUDENTS));
  const invAll = rowsToObjects(sheet(SHEETS.INVENTORY));
  const listingsAll = rowsToObjects(sheet(SHEETS.LISTINGS)).filter(l => l.qty > 0);
  const shopsMap = getAllShopsMap();
  const priceMap = getSupplierPriceMap();
  const txAll = rowsToObjects(sheet(SHEETS.TRANSACTIONS));

  const players = students.map(s => {
    const myInv = invAll.filter(i => i.studentId === s.id && i.qty > 0);
    const invValue = myInv.reduce((sum, i) => sum + (i.qty * (priceMap[i.itemName] || 0)), 0);
    const myListingsValue = listingsAll.filter(l => l.sellerId === s.id)
      .reduce((sum, l) => sum + (l.qty * l.price), 0);
    const shop = shopsMap[s.id];
    return {
      name: s.name,
      cash: Math.round(s.cash),
      netWorth: Math.round(s.cash + invValue + myListingsValue),
      itemsHeld: myInv.reduce((sum, i) => sum + i.qty, 0),
      shopOpen: isShopOpen(shop),
      shopName: shop ? shop.shopName : '',
      joinedAt: s.joinedAt
    };
  }).sort((a, b) => b.netWorth - a.netWorth);

  // ยอดขายรวมและจำนวนธุรกรรม
  let totalVolume = 0;
  txAll.forEach(t => {
    const m = String(t.qtyPriceTotal).match(/=\s*([\d.]+)\s*$/);
    if (m) totalVolume += Number(m[1]);
  });

  // ยอดขายแยกตามสินค้า (จำนวนที่ขายได้ทั้งหมด)
  const itemVolume = {};
  txAll.forEach(t => {
    itemVolume[t.itemName] = (itemVolume[t.itemName] || 0) + 1;
  });
  const topItems = Object.entries(itemVolume)
    .map(([itemName, count]) => ({ itemName, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 6);

  const recentTx = txAll.slice(-25).reverse().map(t => ({
    time: t.time, buyerName: t.buyerName, sellerName: t.sellerName,
    itemName: t.itemName, detail: t.qtyPriceTotal
  }));

  let timeLeftSeconds = null;
  if (cfg.GAME_STARTED_AT) {
    const started = new Date(cfg.GAME_STARTED_AT).getTime();
    const elapsedSec = (Date.now() - started) / 1000;
    timeLeftSeconds = Math.max(0, Math.round(cfg.ROUND_MINUTES * 60 - elapsedSec));
  }

  return {
    gameName: cfg.GAME_NAME,
    joinCode: cfg.JOIN_CODE,
    gameStarted: !!cfg.GAME_STARTED_AT,
    gameEnded: isOpenValue(cfg.GAME_ENDED) || (timeLeftSeconds !== null && timeLeftSeconds <= 0),
    timeLeftSeconds,
    playerCount: students.length,
    shopsOpenCount: Object.values(shopsMap).filter(isShopOpen).length,
    totalTransactions: txAll.length,
    totalVolume: Math.round(totalVolume),
    supplierItems: getSupplierItems(),
    topItems,
    players,
    recentTx
  };
}

function handleTeacherStart(body) {
  if (String(body.key || '') !== getTeacherKey()) return { error: 'รหัสครูไม่ถูกต้อง' };
  startGameCore();
  return handleTeacherState(body);
}

function handleTeacherEnd(body) {
  if (String(body.key || '') !== getTeacherKey()) return { error: 'รหัสครูไม่ถูกต้อง' };
  endGameCore();
  return handleTeacherState(body);
}

function handleTeacherReset(body) {
  if (String(body.key || '') !== getTeacherKey()) return { error: 'รหัสครูไม่ถูกต้อง' };
  if (String(body.confirm || '') !== 'RESET') return { error: 'ต้องยืนยันการรีเซ็ตก่อน' };
  setupCore();
  // หมายเหตุ: setupCore() ตั้ง TEACHER_KEY กลับเป็นค่าเริ่มต้น (teacher123) เสมอ
  // จึงไม่เรียก handleTeacherState ต่อ เพราะ key เดิมที่ใช้ล็อกอินจะไม่ตรงกับ key ใหม่แล้ว
  return { ok: true, reset: true, message: 'รีเซ็ตข้อมูลเรียบร้อยแล้ว รหัสครูถูกตั้งกลับเป็นค่าเริ่มต้น' };
}

function handleTeacherGetConfig(body) {
  if (String(body.key || '') !== getTeacherKey()) return { error: 'รหัสครูไม่ถูกต้อง' };
  const cfg = getConfigMap();
  return {
    ok: true,
    config: {
      GAME_NAME: cfg.GAME_NAME || '',
      JOIN_CODE: cfg.JOIN_CODE || '',
      START_CASH: Number(cfg.START_CASH) || 0,
      ROUND_MINUTES: Number(cfg.ROUND_MINUTES) || 0,
      TRUCK_COST: getTruckCost(),
      TEACHER_KEY: getTeacherKey()
    }
  };
}

function handleTeacherUpdateConfig(body) {
  if (String(body.key || '') !== getTeacherKey()) return { error: 'รหัสครูไม่ถูกต้อง' };
  const c = body.config || {};

  const gameName = String(c.GAME_NAME || '').trim();
  if (!gameName) return { error: 'กรุณาตั้งชื่อเกม' };

  const joinCode = String(c.JOIN_CODE || '').trim().toUpperCase();
  if (!joinCode) return { error: 'กรุณาตั้งรหัสห้อง' };

  const startCash = Number(c.START_CASH);
  if (isNaN(startCash) || startCash < 0) return { error: 'เงินเริ่มต้นต้องเป็นตัวเลขไม่ติดลบ' };

  const roundMinutes = Number(c.ROUND_MINUTES);
  if (isNaN(roundMinutes) || roundMinutes <= 0) return { error: 'เวลาต่อรอบต้องมากกว่า 0 นาที' };

  const truckCost = Number(c.TRUCK_COST);
  if (isNaN(truckCost) || truckCost < 0) return { error: 'ค่าเปิดร้านรถขายของต้องเป็นตัวเลขไม่ติดลบ' };

  const teacherKey = String(c.TEACHER_KEY || '').trim();
  if (!teacherKey || teacherKey.length < 4) return { error: 'รหัสครูต้องมีอย่างน้อย 4 ตัวอักษร' };

  setConfig('GAME_NAME', gameName);
  setConfig('JOIN_CODE', joinCode);
  setConfig('START_CASH', startCash);
  setConfig('ROUND_MINUTES', roundMinutes);
  setConfig('TRUCK_COST', truckCost);
  setConfig('TEACHER_KEY', teacherKey);

  return {
    ok: true,
    config: {
      GAME_NAME: gameName,
      JOIN_CODE: joinCode,
      START_CASH: startCash,
      ROUND_MINUTES: roundMinutes,
      TRUCK_COST: truckCost,
      TEACHER_KEY: teacherKey
    }
  };
}

function handleBuyFromSupplier(body) {
  const student = findStudent(body.studentId);
  if (!student) return { error: 'ไม่พบผู้เล่น' };
  const qty = Number(body.qty);
  if (!qty || qty <= 0) return { error: 'จำนวนไม่ถูกต้อง' };

  const priceMap = getSupplierPriceMap();
  const unitPrice = priceMap[body.itemName];
  if (!unitPrice) return { error: 'ไม่พบสินค้านี้' };

  const total = unitPrice * qty;
  if (student.cash < total) return { error: 'เงินไม่พอซื้อ' };

  updateStudentCash(student.id, student.cash - total);
  addInventory(student.id, body.itemName, qty, unitPrice);
  return { ok: true, spent: total, unitPrice };
}

function handleOpenShop(body) {
  const student = findStudent(body.studentId);
  if (!student) return { error: 'ไม่พบผู้เล่น' };

  const shopName = String(body.shopName || '').trim();
  if (!shopName) return { error: 'ตั้งชื่อร้านด้วยนะ' };
  if (shopName.length > 24) return { error: 'ชื่อร้านยาวไปหน่อย (ไม่เกิน 24 ตัวอักษร)' };
  const tagline = String(body.tagline || '').trim();
  if (tagline.length > 40) return { error: 'คำโปรยร้านยาวไปหน่อย (ไม่เกิน 40 ตัวอักษร)' };

  const style = getTruckStyle(body.styleId);
  const existing = getShop(student.id);
  if (isShopOpen(existing)) return { error: 'เธอเปิดร้านไปแล้วนะ' };

  const cost = getTruckCost();
  if (student.cash < cost) return { error: 'เงินไม่พอซื้อรถขายของ (ต้องมี ' + cost + ' บ.)' };

  updateStudentCash(student.id, student.cash - cost);
  saveShop(student.id, shopName, style.id, true, tagline);
  return { ok: true, spent: cost, shop: { name: shopName, styleId: style.id, tagline, opened: true } };
}

function handleUpdateShop(body) {
  const student = findStudent(body.studentId);
  if (!student) return { error: 'ไม่พบผู้เล่น' };

  const existing = getShop(student.id);
  if (!isShopOpen(existing)) return { error: 'ต้องเปิดร้านก่อนถึงจะแก้ไขได้' };

  const shopName = String(body.shopName || '').trim() || existing.shopName;
  if (shopName.length > 24) return { error: 'ชื่อร้านยาวไปหน่อย (ไม่เกิน 24 ตัวอักษร)' };
  const tagline = body.tagline !== undefined ? String(body.tagline || '').trim() : (existing.tagline || '');
  if (tagline.length > 40) return { error: 'คำโปรยร้านยาวไปหน่อย (ไม่เกิน 40 ตัวอักษร)' };
  const style = getTruckStyle(body.styleId || existing.styleId);

  saveShop(student.id, shopName, style.id, true, tagline);
  return { ok: true, shop: { name: shopName, styleId: style.id, tagline, opened: true } };
}

function handleListForSale(body) {
  const student = findStudent(body.studentId);
  if (!student) return { error: 'ไม่พบผู้เล่น' };
  const qty = Number(body.qty);
  const price = Number(body.price);
  if (!qty || qty <= 0 || !price || price <= 0) return { error: 'ข้อมูลไม่ถูกต้อง' };

  const shop = getShop(student.id);
  if (!isShopOpen(shop)) return { error: 'ต้องเปิดร้านรถขายของก่อนถึงจะวางขายได้ (ดูการ์ด "รถขายของของฉัน")' };

  removeInventory(student.id, body.itemName, qty); // จะ throw ถ้าของไม่พอ
  const sh = sheet(SHEETS.LISTINGS);
  const id = genId('lst');
  sh.appendRow([id, student.id, student.name, body.itemName, qty, price, new Date().toISOString()]);
  return { ok: true, listingId: id };
}

function handleCancelListing(body) {
  const sh = sheet(SHEETS.LISTINGS);
  const rows = rowsToObjects(sh);
  const listing = rows.find(l => l.id === body.listingId && l.sellerId === body.studentId);
  if (!listing) return { error: 'ไม่พบรายการ' };

  addInventory(listing.sellerId, listing.itemName, listing.qty, 0); // คืนของ (ต้นทุนเดิมไม่กระทบเพราะคืนแบบ 0 จะเฉลี่ยลง - ปรับปรุงได้ตามต้องการ)
  sh.getRange(listing._row, 5).setValue(0); // qty = 0 (ปิดรายการ)
  return { ok: true };
}

function handleBuyListing(body) {
  const buyer = findStudent(body.studentId);
  if (!buyer) return { error: 'ไม่พบผู้เล่น' };
  const qty = Number(body.qty);
  if (!qty || qty <= 0) return { error: 'จำนวนไม่ถูกต้อง' };

  const sh = sheet(SHEETS.LISTINGS);
  const rows = rowsToObjects(sh);
  const listing = rows.find(l => l.id === body.listingId);
  if (!listing || listing.qty <= 0) return { error: 'รายการนี้ถูกซื้อ/ยกเลิกไปแล้ว' };
  if (listing.sellerId === buyer.id) return { error: 'ซื้อของร้านตัวเองไม่ได้' };
  if (qty > listing.qty) return { error: 'สินค้าเหลือไม่พอ' };

  const total = listing.price * qty;
  if (buyer.cash < total) return { error: 'เงินไม่พอซื้อ' };

  const seller = findStudent(listing.sellerId);

  // ตัดเงินผู้ซื้อ, เพิ่มของให้ผู้ซื้อ
  updateStudentCash(buyer.id, buyer.cash - total);
  addInventory(buyer.id, listing.itemName, qty, listing.price);

  // เพิ่มเงินผู้ขาย
  updateStudentCash(seller.id, seller.cash + total);

  // อัปเดตจำนวนใน listing
  sh.getRange(listing._row, 5).setValue(listing.qty - qty);

  // บันทึกธุรกรรม
  const tx = sheet(SHEETS.TRANSACTIONS);
  tx.appendRow([genId('tx'), new Date().toISOString(), buyer.id, buyer.name, seller.id, seller.name, listing.itemName, qty + ' x ' + listing.price + ' = ' + total]);

  return { ok: true, total };
}
