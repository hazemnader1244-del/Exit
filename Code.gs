/**
 * Stock Site — Backend (Google Apps Script Web App)
 * الشيت هو الداتابيز:
 *   My user   → اليوزرز (Username | Password | Name | Role | Active)
 *   Products  → المنتجات والمخزون
 *   Orders    → الطلبات (كل طلب بيخصم من المخزون، والإلغاء بيرجّعه)
 *   Movements → سجل كل حركة (وارد / منصرف / جرد / بيع / مرتجع)
 *
 * Deploy: Deploy ← New deployment ← Web app ← Execute as: Me ← Who has access: Anyone
 */

const SH = { USERS: 'My user', PRODUCTS: 'Products', ORDERS: 'Orders', MOVES: 'Movements' };
const HEAD = {
  'My user':   ['Username', 'Password', 'Name', 'Role', 'Active'],
  'Products':  ['ID', 'SKU', 'Name', 'Category', 'Unit', 'Price', 'Cost', 'Stock', 'Min Stock', 'Active', 'Updated'],
  'Orders':    ['Order #', 'Date', 'Customer', 'Phone', 'Items', 'Total', 'Status', 'By', 'Note', 'Cancelled By', 'Cancelled At'],
  'Movements': ['Date', 'Product ID', 'SKU', 'Product', 'Type', 'Qty', 'Before', 'After', 'Ref', 'By', 'Note'],
};
const SESSION_TTL = 21600;   // 6 ساعات (وبتتجدد مع كل استخدام)
const MAX_TRIES = 5;         // محاولات دخول غلط
const LOCK_SEC = 900;        // قفل ربع ساعة
const LIST_LIMIT = 400;      // آخر كام طلب/حركة بيتبعتوا للموقع

/* ======================= ENTRY ======================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📦 Stock Site').addItem('جهّز التابات', 'setup').addToUi();
}

/** شغّلها مرة واحدة: بتعمل التابات الناقصة + يوزر admin مبدئي */
function setup() {
  ensureSheets_();
  const u = sheet_(SH.USERS);
  if (u.getLastRow() < 2) u.appendRow(['admin', 'change-me', 'Admin', 'admin', 'Yes']);
  try { SpreadsheetApp.getUi().alert('✅ التابات جاهزة.\nغيّر باسورد admin من تاب My user، وبعدين Deploy كـ Web app.'); } catch (e) {}
}

function doGet() {
  return json_({ ok: true, app: 'stock-site' });
}

function doPost(e) {
  let q;
  try { q = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'bad_request' }); }
  try {
    ensureSheets_();
    if (q.action === 'login') return json_(login_(q));
    const u = auth_(q.token);
    if (!u) return json_({ ok: false, error: 'unauthorized' });
    const fn = ACTIONS[q.action];
    if (!fn) return json_({ ok: false, error: 'unknown_action' });
    return json_(fn(q, u));
  } catch (err) {
    return json_({ ok: false, error: err.code || 'server', message: String(err.message || err) });
  }
}

const ACTIONS = {
  all: (q, u) => ({ ok: true, data: all_(u) }),

  saveProduct: (q, u) => locked_(() => {
    admin_(u);
    const p = q.product || {};
    const name = String(p.name || '').trim();
    if (!name) throw E('bad_value', 'name');
    const sku = String(p.sku || '').trim();
    const list = products_();
    if (sku && list.some(x => x.active && x.id !== p.id && x.sku.toLowerCase() === sku.toLowerCase())) throw E('dup_sku', sku);

    const sh = sheet_(SH.PRODUCTS);
    const vals = [txt_(sku), name, String(p.cat || '').trim(), String(p.unit || '').trim(), num_(p.price), num_(p.cost)];
    if (p.id) {
      const x = list.find(y => y.id === p.id);
      if (!x) throw E('not_found');
      sh.getRange(x.r, 2, 1, 6).setValues([vals]);
      sh.getRange(x.r, 9, 1, 3).setValues([[num_(p.min), 'Yes', new Date()]]);
    } else {
      const id = nextId_(list);
      const stock = Math.max(0, num_(p.stock));
      sh.appendRow([id].concat(vals, [stock, num_(p.min), 'Yes', new Date()]));
      if (stock) addMove_({ id, sku, name }, 'IN', stock, 0, stock, 'Opening', u, 'رصيد افتتاحي');
    }
    return { ok: true, data: all_(u) };
  }),

  delProduct: (q, u) => locked_(() => {
    admin_(u);
    const x = findP_(q.id);
    sheet_(SH.PRODUCTS).getRange(x.r, 10, 1, 2).setValues([['No', new Date()]]);
    return { ok: true, data: all_(u) };
  }),

  /** type: IN (وارد) | OUT (منصرف) | ADJUST (جرد = الكمية الفعلية) */
  move: (q, u) => locked_(() => {
    const x = findP_(q.id);
    const qty = num_(q.qty);
    const type = String(q.type);
    if (!['IN', 'OUT', 'ADJUST'].includes(type)) throw E('bad_value', 'type');
    if (type === 'ADJUST' ? qty < 0 : qty <= 0) throw E('bad_value', 'qty');
    const after = type === 'IN' ? x.stock + qty : type === 'OUT' ? x.stock - qty : qty;
    if (after < 0) throw E('no_stock', x.name);
    if (after !== x.stock) {
      setStock_(x, after);
      addMove_(x, type, after - x.stock, x.stock, after, '', u, q.note);
    }
    return { ok: true, data: all_(u) };
  }),

  order: (q, u) => locked_(() => {
    const items = (q.items || []).filter(i => num_(i.qty) > 0);
    if (!items.length) throw E('empty_order');
    const byId = {};
    products_().forEach(p => { byId[p.id] = p; });

    const need = {};
    items.forEach(i => {
      const p = byId[i.id];
      if (!p || !p.active) throw E('not_found');
      need[i.id] = (need[i.id] || 0) + num_(i.qty);
    });
    Object.keys(need).forEach(id => {
      if (byId[id].stock < need[id]) throw E('no_stock', byId[id].name + ' (' + byId[id].stock + ')');
    });

    const no = nextOrderNo_();
    const lines = items.map(i => {
      const p = byId[i.id];
      const price = (i.price === '' || i.price == null) ? p.price : num_(i.price);
      const qty = num_(i.qty);
      return { id: p.id, sku: p.sku, name: p.name, qty, price, sum: r2_(price * qty) };
    });
    const total = r2_(lines.reduce((s, l) => s + l.sum, 0));

    Object.keys(need).forEach(id => {
      const p = byId[id];
      const after = p.stock - need[id];
      setStock_(p, after);
      addMove_(p, 'SALE', -need[id], p.stock, after, '#' + no, u, q.customer);
    });
    sheet_(SH.ORDERS).appendRow([no, new Date(), String(q.customer || '').trim(), txt_(q.phone),
      JSON.stringify(lines), total, 'Done', u.name, String(q.note || '').trim(), '', '']);
    return { ok: true, no, data: all_(u) };
  }),

  cancelOrder: (q, u) => locked_(() => {
    admin_(u);
    const o = orders_().find(x => String(x.no) === String(q.no));
    if (!o) throw E('not_found');
    if (o.status !== 'Done') throw E('already_cancelled');
    const byId = {};
    products_().forEach(p => { byId[p.id] = p; });
    o.items.forEach(l => {
      const p = byId[l.id];
      if (!p) return;
      const after = p.stock + l.qty;
      setStock_(p, after);
      addMove_(p, 'RETURN', l.qty, p.stock, after, '#' + o.no, u, 'إلغاء طلب');
      p.stock = after;
    });
    sheet_(SH.ORDERS).getRange(o.r, 7).setValue('Cancelled');
    sheet_(SH.ORDERS).getRange(o.r, 10, 1, 2).setValues([[u.name, new Date()]]);
    return { ok: true, data: all_(u) };
  }),
};

/* ======================= AUTH ======================= */

function login_(q) {
  const name = String(q.name || '').trim().toLowerCase();
  const pass = String(q.pass || '');
  if (!name || !pass) throw E('need');
  const c = CacheService.getScriptCache();
  const k = 'try:' + name;
  const tries = Number(c.get(k) || 0);
  if (tries >= MAX_TRIES) throw E('locked');

  const rows = values_(sheet_(SH.USERS), HEAD[SH.USERS].length);
  if (!rows.length) throw E('no_users_sheet');
  const row = rows.find(r =>
    String(r[0]).trim().toLowerCase() === name &&
    String(r[1]) === pass &&
    !/^(no|false|0|لا)$/i.test(String(r[4]).trim()));
  if (!row) { c.put(k, String(tries + 1), LOCK_SEC); throw E('bad_login'); }
  c.remove(k);

  const user = {
    u: String(row[0]).trim(),
    name: String(row[2] || row[0]).trim(),
    role: /^admin$/i.test(String(row[3]).trim()) ? 'admin' : 'staff',
  };
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  c.put('tok:' + token, JSON.stringify(user), SESSION_TTL);
  return { ok: true, token, user, data: all_(user) };
}

function auth_(token) {
  if (!token) return null;
  const c = CacheService.getScriptCache();
  const s = c.get('tok:' + token);
  if (!s) return null;
  c.put('tok:' + token, s, SESSION_TTL);
  return JSON.parse(s);
}

function admin_(u) { if (u.role !== 'admin') throw E('forbidden'); }

/* ======================= DATA ======================= */

function all_(u) {
  const products = products_().map(p => {
    const o = Object.assign({}, p);
    delete o.r;
    if (u.role !== 'admin') delete o.cost;
    return o;
  });
  const strip = o => { const x = Object.assign({}, o); delete x.r; return x; };
  return {
    user: u,
    products,
    orders: orders_().slice(-LIST_LIMIT).map(strip),
    moves: moves_().slice(-LIST_LIMIT),
  };
}

function products_() {
  return values_(sheet_(SH.PRODUCTS), HEAD[SH.PRODUCTS].length).map((r, i) => ({
    r: i + 2, id: String(r[0]), sku: String(r[1]), name: String(r[2]), cat: String(r[3]), unit: String(r[4]),
    price: num_(r[5]), cost: num_(r[6]), stock: num_(r[7]), min: num_(r[8]),
    active: !/^(no|false|0)$/i.test(String(r[9]).trim()), updated: iso_(r[10]),
  })).filter(p => p.id);
}

function orders_() {
  return values_(sheet_(SH.ORDERS), HEAD[SH.ORDERS].length).map((r, i) => {
    let items = [];
    try { items = JSON.parse(r[4] || '[]'); } catch (e) {}
    return {
      r: i + 2, no: num_(r[0]), date: iso_(r[1]), customer: String(r[2]), phone: String(r[3]), items,
      total: num_(r[5]), status: String(r[6]) || 'Done', by: String(r[7]), note: String(r[8]),
      cby: String(r[9]), cat: iso_(r[10]),
    };
  }).filter(o => o.no);
}

function moves_() {
  return values_(sheet_(SH.MOVES), HEAD[SH.MOVES].length).map(r => ({
    date: iso_(r[0]), pid: String(r[1]), sku: String(r[2]), name: String(r[3]), type: String(r[4]),
    qty: num_(r[5]), before: num_(r[6]), after: num_(r[7]), ref: String(r[8]), by: String(r[9]), note: String(r[10]),
  }));
}

function findP_(id) {
  const x = products_().find(p => p.id === String(id) && p.active);
  if (!x) throw E('not_found');
  return x;
}

function setStock_(p, v) {
  sheet_(SH.PRODUCTS).getRange(p.r, 8).setValue(v);
  sheet_(SH.PRODUCTS).getRange(p.r, 11).setValue(new Date());
}

function addMove_(p, type, qty, before, after, ref, u, note) {
  sheet_(SH.MOVES).appendRow([new Date(), p.id, txt_(p.sku), p.name, type, qty, before, after, ref || '', u.name, String(note || '').trim()]);
}

function nextId_(list) {
  const max = list.reduce((m, p) => Math.max(m, Number(String(p.id).replace(/\D/g, '')) || 0), 0);
  return 'P-' + String(max + 1).padStart(4, '0');
}

function nextOrderNo_() {
  return orders_().reduce((m, o) => Math.max(m, o.no), 1000) + 1;
}

/* ======================= SHEETS / UTILS ======================= */

function ensureSheets_() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEAD).forEach(name => {
    if (ss.getSheetByName(name)) return;
    const sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]])
      .setFontWeight('bold').setBackground('#1c1a19').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (name === SH.USERS) sh.getRange('A:B').setNumberFormat('@');
  });
}

function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function values_(sh, width) {
  const last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, width).getValues();
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { SpreadsheetApp.flush(); lock.releaseLock(); }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function E(code, msg) { const e = new Error(msg || code); e.code = code; return e; }
function num_(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function r2_(n) { return Math.round(n * 100) / 100; }
function iso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }
/** يحافظ على الأصفار في الأول (تليفون / SKU) */
function txt_(v) { v = String(v == null ? '' : v).trim(); return v ? "'" + v : ''; }
