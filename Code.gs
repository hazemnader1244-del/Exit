/**
 * EXIT — Stock & Sales System · Backend (Google Apps Script Web App)
 *
 * الشيت هو الداتابيز:
 *   user / My user → اليوزرز: name | pass | role (admin / employee) | active
 *   Products       → المنتجات والمخزون (كل منتج جديد بياخد كود تلقائي EX00001 …)
 *   Orders         → الطلبات: الإجمالي، الخصم، المدفوع، طريقة الدفع، التكلفة
 *   Payments       → التحصيلات (كل دفعة سطر، والمرتجع بالسالب)
 *   Discounts      → أكواد الخصم: Code | Type (percent / fixed) | Value | Active | Ends | Max Uses | Used
 *   Movements      → سجل كل حركة مخزن
 *
 * Deploy: Deploy ← New deployment ← Web app ← Execute as: Me ← Who has access: Anyone
 * بعد أي تعديل: Deploy ← Manage deployments ← Edit ← New version
 */

const SH = { USERS: 'My user', PRODUCTS: 'Products', ORDERS: 'Orders', PAY: 'Payments', DISC: 'Discounts', MOVES: 'Movements' };
const HEAD = {
  'My user':   ['name', 'pass', 'role', 'active'],
  'Products':  ['ID', 'SKU', 'Name', 'Category', 'Unit', 'Price', 'Cost', 'Stock', 'Min Stock', 'Active', 'Updated'],
  'Orders':    ['Order #', 'Date', 'Customer', 'Phone', 'Items', 'Total', 'Status', 'By', 'Note', 'Cancelled By', 'Cancelled At',
                'Subtotal', 'Discount', 'Paid', 'Method', 'Cost', 'Codes'],
  'Payments':  ['Date', 'Order #', 'Amount', 'Method', 'By', 'Note'],
  'Discounts': ['Code', 'Type', 'Value', 'Active', 'Ends', 'Max Uses', 'Used', 'Note', 'Created By', 'Created At'],
  'Movements': ['Date', 'Product ID', 'SKU', 'Product', 'Type', 'Qty', 'Before', 'After', 'Ref', 'By', 'Note'],
};
const METHODS = ['Cash', 'InstaPay', 'Vodafone Cash', 'Card', 'Other'];
const SESSION_TTL = 21600;   // 6 ساعات (بتتجدد مع الاستخدام)
const MAX_TRIES = 5;         // محاولات دخول غلط
const LOCK_SEC = 900;        // قفل ربع ساعة
const ORDER_LIMIT = 3000;    // آخر كام طلب/دفعة بيتبعتوا للداشبورد
const MOVE_LIMIT = 500;
const SKU_PREFIX = 'EX';     // الأكواد التلقائية: EX00001, EX00002 …

/* ======================= ENTRY ======================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📦 EXIT').addItem('جهّز التابات', 'setup').addToUi();
}

function setup() {
  ensureSheets_();
  const u = usersTab_();
  if (u.getLastRow() < 2 && u.getName() === SH.USERS) u.appendRow(['admin', 'change-me', 'admin', 'Yes']);
  try { SpreadsheetApp.getUi().alert('✅ التابات جاهزة. اعمل Deploy كـ Web app.'); } catch (e) {}
}

function doGet() { return json_({ ok: true, app: 'exit-stock' }); }

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

  /* ---------- Products ---------- */
  saveProduct: (q, u) => locked_(() => {
    admin_(u);
    const p = q.product || {};
    const name = String(p.name || '').trim();
    if (!name) throw E('bad_value', 'name');
    const list = products_();
    let sku = String(p.sku || '').trim();
    if (!sku && !p.id) sku = nextSku_(list);
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

  genCode: (q, u) => locked_(() => {
    admin_(u);
    const list = products_();
    const ids = q.ids && q.ids.length ? q.ids.map(String) : null;
    const sh = sheet_(SH.PRODUCTS);
    let n = 0;
    list.filter(p => p.active && !p.sku && (!ids || ids.includes(p.id))).forEach(p => {
      p.sku = nextSku_(list);
      sh.getRange(p.r, 2).setValue(txt_(p.sku));
      sh.getRange(p.r, 11).setValue(new Date());
      n++;
    });
    return { ok: true, n, data: all_(u) };
  }),

  /** type: IN (وارد) | OUT (منصرف) | ADJUST (جرد) */
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

  /* ---------- Discounts ---------- */
  checkDiscount: (q, u) => {
    const d = findDisc_(discounts_(), q.code);
    return { ok: true, discount: { code: d.code, type: d.type, value: d.value }, amount: calcDisc_(d, num_(q.price), num_(q.qty) || 1) };
  },

  saveDiscount: (q, u) => locked_(() => {
    admin_(u);
    const d = q.discount || {};
    const code = String(d.code || '').trim().toUpperCase().replace(/\s+/g, '');
    if (!/^[A-Z0-9_-]{2,30}$/.test(code)) throw E('bad_value', 'code');
    const type = d.type === 'fixed' ? 'fixed' : 'percent';
    const value = num_(d.value);
    if (value <= 0 || (type === 'percent' && value > 100)) throw E('bad_value', 'value');
    const list = discounts_();
    const orig = String(d.orig || '').toUpperCase();
    if (list.some(x => x.code === code && x.code !== orig)) throw E('dup_code', code);
    const sh = sheet_(SH.DISC);
    const ends = d.ends ? new Date(d.ends + 'T23:59:59') : '';
    const row = [code, type, value, d.active === false ? 'No' : 'Yes', ends, num_(d.max) || '', null, String(d.note || '').trim()];
    const x = orig ? list.find(y => y.code === orig) : null;
    if (x) {
      row[6] = x.used;
      sh.getRange(x.r, 1, 1, 8).setValues([row]);
    } else {
      row[6] = 0;
      sh.appendRow(row.concat([u.name, new Date()]));
    }
    return { ok: true, data: all_(u) };
  }),

  toggleDiscount: (q, u) => locked_(() => {
    admin_(u);
    const x = discounts_().find(d => d.code === String(q.code).toUpperCase());
    if (!x) throw E('bad_code');
    sheet_(SH.DISC).getRange(x.r, 4).setValue(x.active ? 'No' : 'Yes');
    return { ok: true, data: all_(u) };
  }),

  /* ---------- Orders ---------- */
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

    const discs = discounts_();
    const used = {};
    const lines = items.map(i => {
      const p = byId[i.id];
      const qty = num_(i.qty);
      const price = (i.price === '' || i.price == null) ? p.price : num_(i.price);
      if (price < 0) throw E('bad_value', 'price');
      let code = String(i.code || '').trim().toUpperCase(), disc = 0;
      if (code) { const d = findDisc_(discs, code); disc = calcDisc_(d, price, qty); used[d.code] = d; code = d.code; }
      const gross = r2_(price * qty);
      return { id: p.id, sku: p.sku, name: p.name, qty, price, code, disc, sum: r2_(gross - disc), cost: r2_(p.cost * qty) };
    });
    const subtotal = r2_(lines.reduce((s, l) => s + l.price * l.qty, 0));
    const discount = r2_(lines.reduce((s, l) => s + l.disc, 0));
    const total = r2_(subtotal - discount);
    const cost = r2_(lines.reduce((s, l) => s + l.cost, 0));
    const method = METHODS.includes(q.method) ? q.method : 'Cash';
    const paid = (q.paid === '' || q.paid == null) ? total : r2_(Math.max(0, Math.min(total, num_(q.paid))));

    const no = nextOrderNo_();
    Object.keys(need).forEach(id => {
      const p = byId[id];
      const after = p.stock - need[id];
      setStock_(p, after);
      addMove_(p, 'SALE', -need[id], p.stock, after, '#' + no, u, q.customer);
    });
    sheet_(SH.ORDERS).appendRow([no, new Date(), String(q.customer || '').trim(), txt_(q.phone),
      JSON.stringify(lines), total, 'Done', u.name, String(q.note || '').trim(), '', '',
      subtotal, discount, paid, method, cost, Object.keys(used).join(', ')]);
    if (paid > 0) addPay_(no, paid, method, u, 'دفعة الطلب');
    Object.keys(used).forEach(c => sheet_(SH.DISC).getRange(used[c].r, 7).setValue(used[c].used + 1));
    return { ok: true, no, total, discount, paid, data: all_(u) };
  }),

  /** تحصيل جزء أو باقي فلوس طلب آجل */
  pay: (q, u) => locked_(() => {
    admin_(u);
    const o = orders_().find(x => String(x.no) === String(q.no));
    if (!o) throw E('not_found');
    if (o.status !== 'Done') throw E('already_cancelled');
    const amount = r2_(num_(q.amount));
    if (amount <= 0 || amount > r2_(o.total - o.paid) + 0.001) throw E('bad_amount');
    const method = METHODS.includes(q.method) ? q.method : 'Cash';
    sheet_(SH.ORDERS).getRange(o.r, 14).setValue(r2_(o.paid + amount));
    addPay_(o.no, amount, method, u, q.note || 'تحصيل');
    return { ok: true, data: all_(u) };
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
    const sh = sheet_(SH.ORDERS);
    sh.getRange(o.r, 7).setValue('Cancelled');
    sh.getRange(o.r, 10, 1, 2).setValues([[u.name, new Date()]]);
    if (o.paid > 0) { addPay_(o.no, -o.paid, o.method, u, 'استرجاع · إلغاء الطلب'); sh.getRange(o.r, 14).setValue(0); }
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

  const U = usersTab_();
  if (!U || U.getLastRow() < 2) throw E('no_users_sheet');
  const all = U.getRange(1, 1, U.getLastRow(), U.getLastColumn()).getValues();
  const head = all[0].map(h => String(h).trim().toLowerCase());
  const col = names => head.findIndex(h => names.includes(h));
  const cU = col(['username', 'user', 'name', 'اليوزر', 'الاسم']);
  const cP = col(['password', 'pass', 'الباسورد']);
  const cN = col(['display name', 'full name', 'الاسم الكامل']);
  const cR = col(['role', 'الصلاحية']);
  const cA = col(['active', 'مفعل']);
  if (cU < 0 || cP < 0) throw E('no_users_sheet');

  const row = all.slice(1).find(r =>
    String(r[cU]).trim().toLowerCase() === name &&
    String(r[cP]).trim() === pass.trim() &&
    !(cA >= 0 && /^(no|false|0|لا)$/i.test(String(r[cA]).trim())));
  if (!row) { c.put(k, String(tries + 1), LOCK_SEC); throw E('bad_login'); }
  c.remove(k);

  // من غير عمود role أو الخانة فاضية = admin. أي حاجة تانية (employee / staff / موظف) = موظف
  const roleTxt = cR >= 0 ? String(row[cR]).trim() : '';
  const user = {
    u: String(row[cU]).trim(),
    name: String((cN >= 0 && row[cN]) || row[cU]).trim(),
    role: !roleTxt || /^(admin|ادمن|أدمن|مدير)$/i.test(roleTxt) ? 'admin' : 'staff',
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

/** الموظف: مفيش مبيعات ولا تحصيلات ولا أرباح ولا تكلفة */
function all_(u) {
  const admin = u.role === 'admin';
  const strip = o => { const x = Object.assign({}, o); delete x.r; return x; };
  const products = products_().map(p => { const o = strip(p); if (!admin) delete o.cost; return o; });
  let moves = moves_();
  if (!admin) moves = moves.filter(m => m.type !== 'SALE' && m.type !== 'RETURN').map(m => Object.assign({}, m, { note: m.type === 'IN' || m.type === 'OUT' || m.type === 'ADJUST' ? m.note : '' }));
  return {
    user: u,
    products,
    moves: moves.slice(-MOVE_LIMIT),
    orders: admin ? orders_().slice(-ORDER_LIMIT).map(strip) : [],
    payments: admin ? payments_().slice(-ORDER_LIMIT) : [],
    discounts: admin ? discounts_().map(strip) : [],
    methods: METHODS,
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
    const total = num_(r[5]);
    return {
      r: i + 2, no: num_(r[0]), date: iso_(r[1]), customer: String(r[2]), phone: String(r[3]), items,
      total, status: String(r[6]) || 'Done', by: String(r[7]), note: String(r[8]), cby: String(r[9]), cat: iso_(r[10]),
      subtotal: r[11] === '' ? total : num_(r[11]),
      discount: num_(r[12]),
      paid: r[13] === '' ? total : num_(r[13]),            // طلبات قديمة قبل التحديث = مدفوعة
      method: String(r[14] || 'Cash'),
      cost: r[15] === '' ? null : num_(r[15]),
      codes: String(r[16] || ''),
    };
  }).filter(o => o.no);
}

function payments_() {
  return values_(sheet_(SH.PAY), HEAD[SH.PAY].length).map(r => ({
    date: iso_(r[0]), no: num_(r[1]), amount: num_(r[2]), method: String(r[3] || 'Cash'), by: String(r[4]), note: String(r[5]),
  })).filter(p => p.amount);
}

function discounts_() {
  return values_(sheet_(SH.DISC), HEAD[SH.DISC].length).map((r, i) => ({
    r: i + 2, code: String(r[0]).trim().toUpperCase(), type: /fixed|ثابت|جنيه/i.test(String(r[1])) ? 'fixed' : 'percent',
    value: num_(r[2]), active: !/^(no|false|0|لا)$/i.test(String(r[3]).trim()), ends: iso_(r[4]),
    max: num_(r[5]), used: num_(r[6]), note: String(r[7]), by: String(r[8]), created: iso_(r[9]),
  })).filter(d => d.code);
}

function moves_() {
  return values_(sheet_(SH.MOVES), HEAD[SH.MOVES].length).map(r => ({
    date: iso_(r[0]), pid: String(r[1]), sku: String(r[2]), name: String(r[3]), type: String(r[4]),
    qty: num_(r[5]), before: num_(r[6]), after: num_(r[7]), ref: String(r[8]), by: String(r[9]), note: String(r[10]),
  }));
}

function findDisc_(list, code) {
  code = String(code || '').trim().toUpperCase();
  const d = list.find(x => x.code === code);
  if (!d) throw E('bad_code', code);
  if (!d.active) throw E('code_inactive', code);
  if (d.ends) { const end = new Date(d.ends); if (end < new Date()) throw E('code_expired', code); }
  if (d.max > 0 && d.used >= d.max) throw E('code_used_up', code);
  return d;
}

/** percent: % من إجمالي السطر · fixed: مبلغ ثابت من كل قطعة */
function calcDisc_(d, price, qty) {
  const gross = price * qty;
  return r2_(d.type === 'fixed' ? Math.min(d.value * qty, gross) : gross * Math.min(d.value, 100) / 100);
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

function addPay_(no, amount, method, u, note) {
  sheet_(SH.PAY).appendRow([new Date(), no, amount, method, u.name, String(note || '').trim()]);
}

function nextId_(list) {
  const max = list.reduce((m, p) => Math.max(m, Number(String(p.id).replace(/\D/g, '')) || 0), 0);
  return 'P-' + String(max + 1).padStart(4, '0');
}

function nextSku_(list) {
  const re = new RegExp('^' + SKU_PREFIX + '(\\d+)$', 'i');
  const max = list.reduce((m, p) => { const x = re.exec(String(p.sku).trim()); return x ? Math.max(m, Number(x[1])) : m; }, 0);
  return SKU_PREFIX + String(max + 1).padStart(5, '0');
}

function nextOrderNo_() {
  return orders_().reduce((m, o) => Math.max(m, o.no), 1000) + 1;
}

/* ======================= SHEETS / UTILS ======================= */

function ensureSheets_() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEAD).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (name === SH.USERS && !sh && usersTab_()) return;
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.setFrozenRows(1);
      if (name === SH.USERS) sh.getRange('A:B').setNumberFormat('@');
    }
    // أعمدة جديدة بتتضاف في الآخر من غير ما تبوظ الداتا القديمة
    if (name !== SH.USERS && sh.getLastColumn() < HEAD[name].length) {
      sh.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]])
        .setFontWeight('bold').setBackground('#0a0a0a').setFontColor('#ece9e3');
    }
  });
}

function usersTab_() {
  return SpreadsheetApp.getActive().getSheets()
    .find(sh => ['my user', 'user', 'users'].includes(sh.getName().trim().toLowerCase())) || null;
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
function txt_(v) { v = String(v == null ? '' : v).trim(); return v ? "'" + v : ''; }
