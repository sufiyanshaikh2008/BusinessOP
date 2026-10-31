/**
 * BusinessOP — Google Apps Script backend
 * Storage: Google Sheets (1 spreadsheet = 1 company database)
 * Money is stored in PAISE (integers) — same as the client.
 *
 * FILES NEEDED:
 *   Code.gs     (this file)
 *   Index.html  (create an HTML file named exactly "Index")
 *
 * FIRST RUN:
 *   1) Run setup()        -> creates the spreadsheet + tabs
 *   2) (optional) Run seedDemoData()  -> loads sample products/contacts
 *   3) Deploy -> New deployment -> Web app
 *        Execute as: Me
 *        Who has access: Only myself (or Anyone)
 *   4) Open the Web App URL
 */

const APP_NAME = 'BusinessOP';
const PROP_DB  = 'BOP_DB_ID';

/* ---------------------------------------------------------------
   SCHEMA  (sheet name -> header row)
   --------------------------------------------------------------- */
const SCHEMA = {
  Products:     ['id','name','price','cost','stock','min','updated'],
  Customers:    ['id','name','phone','updated'],
  Vendors:      ['id','name','phone','updated'],
  Invoices:     ['id','no','date','cust','total','paid','updated'],
  InvoiceItems: ['invoiceId','productId','name','qty','price','cost'],
  Purchases:    ['id','vendor','productId','name','qty','cost','date'],
  Payments:     ['id','invoiceId','amount','date','mode'],
  Meta:         ['key','value']
};

/* ---------------------------------------------------------------
   WEB APP ENTRY
   --------------------------------------------------------------- */
function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('Index');
  // Escape "<" so the JSON can never terminate the <script> tag.
  tpl.bootstrapJson = JSON.stringify(getBootstrap()).replace(/</g, '\\u003c');
  return tpl.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------------------------------------------------------------
   DATABASE LAYER
   --------------------------------------------------------------- */
function getDb_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(PROP_DB);

  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      // spreadsheet deleted / no access -> fall through and make a new one
    }
  }

  const ss = SpreadsheetApp.create(APP_NAME + ' DB');
  props.setProperty(PROP_DB, ss.getId());
  Object.keys(SCHEMA).forEach(function (n) { ensureSheet_(ss, n); });
  return ss;
}

function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  const head = SCHEMA[name];

  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readRows_(name) {
  const ss = getDb_();
  const sh = ensureSheet_(ss, name);
  const last = sh.getLastRow();
  if (last < 2) return [];

  const head = SCHEMA[name];
  const values = sh.getRange(2, 1, last - 1, head.length).getValues();

  return values
    .filter(function (r) { return String(r[0]).length > 0; })
    .map(function (r) {
      const o = {};
      head.forEach(function (h, i) { o[h] = r[i]; });
      return o;
    });
}

/** Insert or update a row whose first column is a unique key. */
function writeRow_(name, obj, keyCol) {
  const ss = getDb_();
  const sh = ensureSheet_(ss, name);
  const head = SCHEMA[name];
  const key = String(obj[keyCol || head[0]]);
  const row = head.map(function (h) {
    const v = obj[h];
    return (v === undefined || v === null) ? '' : v;
  });

  const last = sh.getLastRow();
  if (last >= 2) {
    const col = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < col.length; i++) {
      if (String(col[i][0]) === key) {
        sh.getRange(i + 2, 1, 1, head.length).setValues([row]);
        return obj;
      }
    }
  }
  sh.appendRow(row);
  return obj;
}

function deleteRow_(name, id) {
  const ss = getDb_();
  const sh = ensureSheet_(ss, name);
  const last = sh.getLastRow();
  if (last < 2) return false;

  const col = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) {
    if (String(col[i][0]) === String(id)) {
      sh.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

function metaGet_(key, def) {
  const r = readRows_('Meta').filter(function (x) { return String(x.key) === String(key); })[0];
  return r ? r.value : def;
}

function metaSet_(key, val) {
  writeRow_('Meta', { key: key, value: val }, 'key');
}

/* ---------------------------------------------------------------
   BOOTSTRAP — everything the client needs on first paint
   --------------------------------------------------------------- */
function getBootstrap() {
  const products = readRows_('Products').map(function (r) {
    return {
      id: String(r.id), name: String(r.name),
      price: Number(r.price) || 0, cost: Number(r.cost) || 0,
      stock: Number(r.stock) || 0, min: Number(r.min) || 0
    };
  });

  const customers = readRows_('Customers').map(function (r) {
    return { id: String(r.id), name: String(r.name), phone: String(r.phone || '') };
  });

  const vendors = readRows_('Vendors').map(function (r) {
    return { id: String(r.id), name: String(r.name), phone: String(r.phone || '') };
  });

  // group invoice items
  const itemsByInv = {};
  readRows_('InvoiceItems').forEach(function (r) {
    const k = String(r.invoiceId);
    (itemsByInv[k] = itemsByInv[k] || []).push({
      id: String(r.productId),
      name: String(r.name),
      qty: Number(r.qty) || 0,
      price: Number(r.price) || 0,
      cost: Number(r.cost) || 0
    });
  });

  const invoices = readRows_('Invoices').map(function (r) {
    return {
      id: String(r.id), no: String(r.no),
      date: Number(r.date) || Date.now(),
      cust: String(r.cust),
      total: Number(r.total) || 0,
      paid: Number(r.paid) || 0,
      items: itemsByInv[String(r.id)] || [],
      pays: []
    };
  });

  const purchases = readRows_('Purchases').map(function (r) {
    return {
      id: String(r.id), vendor: String(r.vendor), pid: String(r.productId),
      name: String(r.name), qty: Number(r.qty) || 0,
      cost: Number(r.cost) || 0, date: Number(r.date) || Date.now()
    };
  });

  return {
    products: products,
    customers: customers,
    vendors: vendors,
    invoices: invoices,
    purchases: purchases,
    seq: Number(metaGet_('invoiceSeq', 0)) || 0
  };
}

/* ---------------------------------------------------------------
   PRODUCTS
   --------------------------------------------------------------- */
function saveProduct(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const rec = {
      id: String(p.id || ('p' + Utilities.getUuid().slice(0, 8))),
      name: String(p.name || '').trim(),
      price: Math.max(0, Math.round(Number(p.price) || 0)),
      cost:  Math.max(0, Math.round(Number(p.cost)  || 0)),
      stock: Math.max(0, Math.round(Number(p.stock) || 0)),
      min:   Math.max(0, Math.round(Number(p.min)   || 0)),
      updated: Date.now()
    };
    if (!rec.name) throw new Error('Product name is required.');

    // duplicate-name guard
    const dup = readRows_('Products').filter(function (r) {
      return String(r.id) !== rec.id &&
             String(r.name).toLowerCase() === rec.name.toLowerCase();
    });
    if (dup.length) throw new Error('A product with this name already exists.');

    writeRow_('Products', rec);
    return { ok: true, product: rec };
  } finally {
    lock.releaseLock();
  }
}

function deleteProduct(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return { ok: deleteRow_('Products', id) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------------------------------------------------
   CONTACTS  (kind = 'cust' | 'vend')
   --------------------------------------------------------------- */
function saveContact(kind, c) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const isVendor = kind === 'vend';
    const sheet = isVendor ? 'Vendors' : 'Customers';
    const prefix = isVendor ? 'v' : 'c';

    const rec = {
      id: String(c.id || (prefix + Utilities.getUuid().slice(0, 8))),
      name: String(c.name || '').trim(),
      phone: String(c.phone || '').replace(/[\s-]/g, ''),
      updated: Date.now()
    };

    if (!rec.name) throw new Error('Name is required.');
    if (rec.phone && !/^\d{10}$/.test(rec.phone)) throw new Error('Phone must be 10 digits.');

    const dup = readRows_(sheet).filter(function (r) {
      return String(r.id) !== rec.id &&
             String(r.name).toLowerCase() === rec.name.toLowerCase();
    });
    if (dup.length) throw new Error('This name already exists.');

    writeRow_(sheet, rec);
    return { ok: true, contact: rec };
  } finally {
    lock.releaseLock();
  }
}

function deleteContact(kind, id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return { ok: deleteRow_(kind === 'vend' ? 'Vendors' : 'Customers', id) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------------------------------------------------
   INVENTORY — receive stock from a vendor
   payload = { vendor, productId, qty, cost }
   --------------------------------------------------------------- */
function receiveStock(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const qty = Math.round(Number(p.qty) || 0);
    const cost = Math.max(0, Math.round(Number(p.cost) || 0));
    if (qty < 1) throw new Error('Quantity must be 1 or more.');

    const prod = readRows_('Products').filter(function (r) {
      return String(r.id) === String(p.productId);
    })[0];
    if (!prod) throw new Error('Product not found.');

    const newStock = (Number(prod.stock) || 0) + qty;

    writeRow_('Products', {
      id: prod.id, name: prod.name,
      price: Number(prod.price) || 0, cost: cost,
      stock: newStock, min: Number(prod.min) || 0,
      updated: Date.now()
    });

    const purchase = {
      id: 'u' + Utilities.getUuid().slice(0, 8),
      vendor: String(p.vendor),
      productId: String(prod.id),
      name: String(prod.name),
      qty: qty, cost: cost, date: Date.now()
    };
    writeRow_('Purchases', purchase);

    return { ok: true, purchase: purchase, stock: newStock, cost: cost };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------------------------------------------------
   SALES — create invoice (atomic: stock + items + payment)
   payload = { cust, items:[{id,name,qty,price,cost}], pm, partPaid }
   --------------------------------------------------------------- */
function createInvoice(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const items = (p.items || []).map(function (it) {
      return {
        id: String(it.id),
        name: String(it.name),
        qty: Math.round(Number(it.qty) || 0),
        price: Math.max(0, Math.round(Number(it.price) || 0)),
        cost: Math.max(0, Math.round(Number(it.cost) || 0))
      };
    }).filter(function (it) { return it.qty > 0; });

    if (!items.length) throw new Error('Invoice has no items.');

    const byId = {};
    readRows_('Products').forEach(function (r) { byId[String(r.id)] = r; });

    // running stock check (handles the same product on 2 lines)
    const after = {};
    items.forEach(function (it) {
      const pr = byId[it.id];
      if (!pr) throw new Error('Product not found: ' + it.name);
      const cur = (after[it.id] !== undefined) ? after[it.id] : (Number(pr.stock) || 0);
      if (cur < it.qty) throw new Error(pr.name + ': only ' + cur + ' in stock.');
      after[it.id] = cur - it.qty;
    });

    const total = items.reduce(function (s, it) { return s + it.qty * it.price; }, 0);
    const cust = String(p.cust || 'walk');

    let paid = total;
    if (p.pm === 'later') paid = 0;
    if (p.pm === 'part') {
      paid = Math.round(Number(p.partPaid) || 0);
      if (!(paid > 0 && paid < total)) {
        throw new Error('Part payment must be above 0 and below the total.');
      }
    }
    if (p.pm !== 'full' && cust === 'walk') {
      throw new Error('Choose a customer for credit sales.');
    }

    const id = 'i' + Utilities.getUuid().slice(0, 8);
    const seq = Number(metaGet_('invoiceSeq', 0)) + 1;
    metaSet_('invoiceSeq', seq);
    const no = 'INV-' + String(seq).padStart(4, '0');
    const date = Date.now();

    const invoice = { id: id, no: no, date: date, cust: cust,
                      total: total, paid: paid, items: items, pays: [] };

    // 1) invoice header
    writeRow_('Invoices', {
      id: id, no: no, date: date, cust: cust,
      total: total, paid: paid, updated: date
    });

    // 2) invoice lines
    const shItems = ensureSheet_(getDb_(), 'InvoiceItems');
    items.forEach(function (it) {
      shItems.appendRow([id, it.id, it.name, it.qty, it.price, it.cost]);
    });

    // 3) payment record
    if (paid > 0) {
      writeRow_('Payments', {
        id: 'pay' + Utilities.getUuid().slice(0, 8),
        invoiceId: id, amount: paid, date: date, mode: 'sale'
      });
    }

    // 4) decrement stock
    items.forEach(function (it) {
      const pr = byId[it.id];
      writeRow_('Products', {
        id: pr.id, name: pr.name,
        price: Number(pr.price) || 0, cost: Number(pr.cost) || 0,
        stock: after[it.id], min: Number(pr.min) || 0, updated: date
      });
    });

    return { ok: true, invoice: invoice, seq: seq };
  } finally {
    lock.releaseLock();
  }
}

/** Record a payment against an existing invoice. */
function recordPayment(invoiceId, amount) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const inv = readRows_('Invoices').filter(function (r) {
      return String(r.id) === String(invoiceId);
    })[0];
    if (!inv) throw new Error('Invoice not found.');

    const total = Number(inv.total) || 0;
    const paid  = Number(inv.paid)  || 0;
    const due   = total - paid;
    const a     = Math.round(Number(amount) || 0);

    if (a <= 0)  throw new Error('Enter an amount above 0.');
    if (a > due) throw new Error('Amount exceeds balance due.');

    const newPaid = paid + a;

    writeRow_('Invoices', {
      id: inv.id, no: inv.no, date: Number(inv.date), cust: inv.cust,
      total: total, paid: newPaid, updated: Date.now()
    });

    writeRow_('Payments', {
      id: 'pay' + Utilities.getUuid().slice(0, 8),
      invoiceId: String(inv.id), amount: a, date: Date.now(), mode: 'settle'
    });

    return { ok: true, paid: newPaid, due: total - newPaid };
  } finally {
    lock.releaseLock();
  }
}

/** Cancel an invoice and put the stock back. */
function deleteInvoice(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const items = readRows_('InvoiceItems').filter(function (r) {
      return String(r.invoiceId) === String(id);
    });

    const byId = {};
    readRows_('Products').forEach(function (r) { byId[String(r.id)] = r; });

    items.forEach(function (it) {
      const pr = byId[String(it.productId)];
      if (!pr) return;
      writeRow_('Products', {
        id: pr.id, name: pr.name,
        price: Number(pr.price) || 0, cost: Number(pr.cost) || 0,
        stock: (Number(pr.stock) || 0) + (Number(it.qty) || 0),
        min: Number(pr.min) || 0, updated: Date.now()
      });
    });

    // wipe lines
    const ss = getDb_();
    const sh = ensureSheet_(ss, 'InvoiceItems');
    for (let i = sh.getLastRow(); i >= 2; i--) {
      if (String(sh.getRange(i, 1).getValue()) === String(id)) sh.deleteRow(i);
    }

    deleteRow_('Invoices', id);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------------------------------------------------
   ADMIN HELPERS — run once from the editor
   --------------------------------------------------------------- */
function setup() {
  const ss = getDb_();
  Object.keys(SCHEMA).forEach(function (n) { ensureSheet_(ss, n); });
  Logger.log('Database ready: ' + ss.getUrl());
}

function getDatabaseUrl() {
  return getDb_().getUrl();
}

/** Load the demo catalogue. Safe to skip. */
function seedDemoData() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (readRows_('Products').length) {
      throw new Error('Products already exist — clear the sheet first.');
    }

    const products = [
      ['Basmati Rice 5kg', 520, 440, 24, 8], ['Tata Salt 1kg', 28, 22, 60, 15],
      ['Amul Butter 100g', 58, 50, 6, 10],  ['Parle-G Biscuits', 10, 8, 120, 30],
      ['Toor Dal 1kg', 165, 140, 14, 10],   ['Sunflower Oil 1L', 148, 128, 4, 8],
      ['Maggi Noodles 4-pack', 56, 46, 40, 12], ['Surf Excel 1kg', 135, 112, 18, 6],
      ['Colgate Paste 100g', 98, 80, 22, 8], ['Tea Powder 250g', 135, 110, 0, 6]
    ];
    products.forEach(function (a, i) {
      writeRow_('Products', {
        id: 'p' + (i + 1), name: a[0],
        price: a[1] * 100, cost: a[2] * 100,
        stock: a[3], min: a[4], updated: Date.now()
      });
    });

    [
      { id: 'c1', name: 'Anita Sharma',  phone: '9825012345' },
      { id: 'c2', name: 'Rajesh Traders', phone: '9898011122' },
      { id: 'c3', name: 'Meera Patel',   phone: '' }
    ].forEach(function (c) {
      writeRow_('Customers', { id: c.id, name: c.name, phone: c.phone, updated: Date.now() });
    });

    [
      { id: 'v1', name: 'Gujarat Wholesale Mart', phone: '9879012345' },
      { id: 'v2', name: 'Shree FMCG Supplies',    phone: '9374055566' }
    ].forEach(function (v) {
      writeRow_('Vendors', { id: v.id, name: v.name, phone: v.phone, updated: Date.now() });
    });

    return 'Demo data seeded.';
  } finally {
    lock.releaseLock();
  }
}

/** Danger zone — deletes every row but keeps the sheets. */
function resetDatabase() {
  const ss = getDb_();
  Object.keys(SCHEMA).forEach(function (n) {
    const sh = ensureSheet_(ss, n);
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, SCHEMA[n].length).clearContent();
    }
  });
  return 'Database cleared.';
}
