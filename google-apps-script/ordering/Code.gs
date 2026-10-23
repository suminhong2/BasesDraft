////////////////////////////////////////////////////////////////////////////
//                   Settings - CHANGE THESE FOR YOUR LAB                 //
////////////////////////////////////////////////////////////////////////////
const CONFIG = {
  ordersSheet: 'Orders',
  settingsSheet: 'Settings',      // grants + budgets, and who places orders
  analysisSheet: 'Analysis',
  menuName: '🧾 Lab orders 🧾',
  labName: 'Wontakal',
  orderIdPrefix: 'ORD-',          // orders become ORD-0001, ORD-0002...
  currency: '$',
  headerRow: 2,                   // row 1 holds the title, row 2 the column names
  lockTimeoutMs: 10000,
  monthsInAnalysis: 12,           // how many recent months "Orders per month" shows

  // Your other MegaBases sheets. Paste each one's whole address from the browser bar.
  // Strain orders check the strain collection; plasmid and oligo are left blank for now.
  strainSheetUrl: 'PASTE_YOUR_STRAIN_SHEET_URL_HERE',
  plasmidSheetUrl: '',            // not built yet
  oligoSheetUrl: '',              // not built yet
};

// What kind of order. Plasmid and Oligo are switched off in the form until those sheets exist.
const ORDER_TYPES = ['Strain', 'Plasmid', 'Oligo', 'Supplies / other'];
const TYPES_READY = ['Strain', 'Supplies / other'];

// Columns of the Orders tab, in order. The script finds columns by these names.
const COLS = [
  'Order ID', 'Request date', 'Requested by', 'Requester email',
  'Item', 'Vendor', 'Catalog #', 'Link', 'Quantity', 'Unit price', 'Total',
  'Grant', 'Needed by', 'Notes',
  'Status', 'Ordered by', 'Order date', 'Confirmation / PO #', 'Received date',
  'Type',        // Strain, Plasmid, Oligo or Supplies / other
  'Linked to',   // filled in when the request comes from the strain collection (e.g. "Strain CB1370")
];
const STATUSES = ['Requested', 'Ordered', 'Received', 'Cancelled'];
const COL = {};
COLS.forEach((name, i) => { COL[name] = i + 1; });

const THEME = {
  header: '#2c5f6e',
  headerText: '#ffffff',
  settings: '#eef4f5',
  settingsText: '#3b4d52',
  stripe: '#f5f9fa',
  grid: '#dde6e8',
};
const STATUS_COLORS = {
  Requested: ['#fff4e5', '#8a5a00'],
  Ordered: ['#e8f0fe', '#1a4fa0'],
  Received: ['#e6f4ea', '#1e6b34'],
  Cancelled: ['#f1f3f4', '#80868b'],
};
const COLUMN_WIDTHS = [85, 95, 120, 170, 240, 120, 100, 70, 70, 80, 85, 120, 90, 220, 95, 120, 95, 130, 95, 110, 130];

////////////////////////////////////////////////////////////////////////////
//                                Utilities                               //
////////////////////////////////////////////////////////////////////////////
const DATE_FORMAT = 'yyyy-mm-dd';

function getSheetOrFail(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error('Could not find a sheet called "' + name + '". Run the menu item "Set up the order sheets" first.');
  }
  return sheet;
}
function getOrders() { return getSheetOrFail(CONFIG.ordersSheet); }
function getSettings() { return getSheetOrFail(CONFIG.settingsSheet); }

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(CONFIG.lockTimeoutMs);
  try {
    const result = fn();
    SpreadsheetApp.flush();
    return result;
  } finally {
    lock.releaseLock();
  }
}

function toDate(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const m = String(v || '').trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}
function formatDay(d) {
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
}
function money(n) {
  return CONFIG.currency + Number(n || 0).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
function toNumber(v) {
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}
function isEmail(e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e); }

// All order rows as objects: { row, 'Order ID': ..., 'Item': ..., ... }
function readOrders() {
  const sheet = getOrders();
  const n = sheet.getLastRow() - CONFIG.headerRow;
  if (n < 1) return [];
  const values = sheet.getRange(CONFIG.headerRow + 1, 1, n, COLS.length).getValues();
  return values
    .map((r, i) => {
      const o = { row: CONFIG.headerRow + 1 + i };
      COLS.forEach((name, c) => { o[name] = r[c]; });
      return o;
    })
    .filter(o => String(o['Order ID']).trim() !== '' || String(o['Item']).trim() !== '');
}

////////////////////////////////////////////////////////////////////////////
//          Settings tab: grants (with budgets) and who places orders     //
////////////////////////////////////////////////////////////////////////////
// Grants in columns A-C (Grant, Budget, Notes), people who order in E-F (Name, Email), from row 3
function getGrants() {
  const s = getSettings();
  const n = Math.max(s.getLastRow() - 2, 0);
  if (!n) return [];
  return s.getRange(3, 1, n, 2).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({ name: String(r[0]).trim(), budget: toNumber(r[1]) }));
}

function getOrderers() {
  const s = getSettings();
  const n = Math.max(s.getLastRow() - 2, 0);
  if (!n) return [];
  return s.getRange(3, 5, n, 2).getValues()
    .filter(r => isEmail(String(r[1]).trim()))
    .map(r => ({ name: String(r[0]).trim(), email: String(r[1]).trim() }));
}

// Menu item: choose the email addresses notified about every new order.
// They are kept in the Settings tab (columns E-F), so orders sent from the strain sheet notify them too.
function setOrderEmails() {
  const ui = SpreadsheetApp.getUi();
  const current = getOrderers();
  const res = ui.prompt('Who gets an email for every new order?',
    'Email addresses, separated by commas. Leave empty to turn order emails off.\nNow: ' +
    (current.length ? current.map(p => p.email).join(', ') : 'nobody'), ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const list = res.getResponseText().split(',').map(e => e.trim()).filter(e => e !== '')
    .filter((e, i, arr) => arr.map(x => x.toLowerCase()).indexOf(e.toLowerCase()) === i);
  const bad = list.filter(e => !isEmail(e));
  if (bad.length) {
    ui.alert('These don\'t look like email addresses: ' + bad.join(', ') + '. Nothing was changed.');
    return;
  }
  withLock(() => {
    const s = getSettings();
    const n = Math.max(s.getLastRow() - 2, 0);
    if (n) s.getRange(3, 5, n, 2).clearContent();
    // Keep the names already written next to addresses that stay on the list
    const rows = list.map(e => {
      const old = current.find(p => p.email.toLowerCase() === e.toLowerCase());
      return [old ? old.name : '', e];
    });
    if (rows.length) s.getRange(3, 5, rows.length, 2).setValues(rows);
  });
  SpreadsheetApp.getActive().toast(list.length ? 'New orders are emailed to: ' + list.join(', ') : 'Order emails are off.');
}

// Used by the request form to fill its dropdowns
function getFormInfo() {
  return {
    types: ORDER_TYPES.map(t => ({ name: t, ready: TYPES_READY.indexOf(t) !== -1 })),
    strainSheet: !!strainSheetUrl(),
    grants: getGrants().map(g => g.name),
    email: Session.getActiveUser().getEmail() || '',
    vendors: Array.from(new Set(readOrders().map(o => String(o['Vendor']).trim()).filter(v => v))).sort(),
  };
}

////////////////////////////////////////////////////////////////////////////
//          1. Someone asks for something: new row + email the orderers   //
////////////////////////////////////////////////////////////////////////////
function startOrderForm() {
  const html = HtmlService.createTemplateFromFile('OrderForm').evaluate().setWidth(720).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Request an order');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function nextOrderId(orders) {
  const re = new RegExp('^' + CONFIG.orderIdPrefix.replace(/[^A-Za-z0-9]/g, '\\$&') + '(\\d+)$');
  let max = 0;
  orders.forEach(o => {
    const m = String(o['Order ID']).trim().match(re);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return CONFIG.orderIdPrefix + String(max + 1).padStart(4, '0');
}

// Called by the form. Returns a message for the form to show.
function submitOrder(data) {
  const type = String(data.type || 'Supplies / other');
  if (ORDER_TYPES.indexOf(type) === -1) throw new Error('Choose what kind of order it is.');
  if (TYPES_READY.indexOf(type) === -1) throw new Error(type + ' orders are not set up yet. Use "Supplies / other" for now.');
  let linked = '';
  if (type === 'Strain') {
    const strain = String(data.strain || '').trim();
    if (!/^[A-Za-z0-9_.-]+$/.test(strain)) throw new Error('Type a strain name like CB1370.');
    const geno = String(data.geno || '').trim();
    data.item = 'C. elegans strain ' + strain + (geno ? ' (' + geno + ')' : '');
    data.catalog = strain;
    data.vendor = data.source || 'CGC';
    if (!data.link) data.link = strainOrderUrl(data.vendor, strain);
    const url = strainSheetUrl();
    if (url) linked = '=HYPERLINK("' + url + '","Strain ' + strain + '")';
  }
  const item = String(data.item || '').trim();
  const name = String(data.name || '').trim();
  const email = String(data.email || '').trim();
  if (!item) throw new Error('Type what you need.');
  if (!name) throw new Error('Type your name.');
  if (!isEmail(email)) throw new Error('Type your email, so you hear back when it is ordered.');
  const qty = toNumber(data.quantity) || 1;
  const price = data.price === '' || data.price == null ? '' : toNumber(data.price);

  const saved = withLock(() => {
    const sheet = getOrders();
    const id = nextOrderId(readOrders());
    const row = {};
    COLS.forEach(c => { row[c] = ''; });
    Object.assign(row, {
      'Order ID': id, 'Request date': new Date(), 'Requested by': name, 'Requester email': email,
      'Item': item, 'Vendor': String(data.vendor || '').trim(), 'Catalog #': String(data.catalog || '').trim(),
      'Link': String(data.link || '').trim(), 'Quantity': qty, 'Unit price': price,
      'Grant': String(data.grant || '').trim(), 'Needed by': toDate(data.needed) || '',
      'Notes': String(data.notes || '').trim(), 'Status': 'Requested', 'Type': type,
    });
    const r = sheet.getLastRow() + 1;
    if (r > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    sheet.getRange(r, 1, 1, COLS.length).setValues([COLS.map(c => row[c])]);
    sheet.getRange(r, COL['Total']).setFormula(totalFormula(r));
    if (linked) sheet.getRange(r, COL['Linked to']).setFormula(linked);
    formatOrderRow(sheet, r);
    row.row = r;
    return row;
  });

  const mail = emailOrderers(saved);
  refreshAnalysisQuietly();
  const msg = saved['Order ID'] + ' saved. ' + mail;
  SpreadsheetApp.getActive().toast(msg, 'Order requested', 8);
  return msg;
}

////////////////////////////////////////////////////////////////////////////
//            Strain orders: check the strain collection and the CGC      //
////////////////////////////////////////////////////////////////////////////
function strainSheetUrl() {
  const url = String(CONFIG.strainSheetUrl || '').trim();
  return !url || /PASTE_YOUR/.test(url) ? '' : url;
}

function strainOrderUrl(source, strain) {
  if (source === 'CGC') return 'https://cgc.umn.edu/strain/' + encodeURIComponent(strain);
  if (source === 'NBRP') return 'https://shigen.nig.ac.jp/c.elegans/';
  return '';
}

// Called by the form's "Check" button. Looks the strain up in your strain collection
// (is it already in a box?) and on the CGC (genotype, can it be ordered there?).
function checkStrain(strainName) {
  const name = String(strainName || '').trim();
  if (!/^[A-Za-z0-9_.-]+$/.test(name)) throw new Error('Type a strain name like CB1370 first.');
  const out = { name: name, collection: null, cgc: null, note: '' };
  const url = strainSheetUrl();
  if (url) {
    try {
      const book = /^https?:/.test(url) ? SpreadsheetApp.openByUrl(url) : SpreadsheetApp.openById(url);
      const sheet = book.getSheetByName('Strains');
      const values = sheet.getRange(1, 1, sheet.getLastRow(), 10).getDisplayValues();
      const row = values.find(r => r[0].trim().toUpperCase() === name.toUpperCase());
      // Columns: A strain number, B alleles / arrays, J box location
      if (row) out.collection = { geno: row[1], box: /box/i.test(row[9]) ? row[9] : '' };
    } catch (e) {
      out.note = 'Could not open the strain collection: ' + e.message;
    }
  } else {
    out.note = 'The strain collection is not connected (strainSheetUrl in Code.gs).';
  }
  try { out.cgc = lookupCgcStrain(name); } catch (e) { console.log('CGC lookup failed: ' + e); }
  return out;
}

function htmlToText(html) {
  return html
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#0?39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

function lookupCgcStrain(name) {
  const response = UrlFetchApp.fetch('https://cgc.umn.edu/strain/' + encodeURIComponent(name),
                                     { muteHttpExceptions: true, followRedirects: true });
  if (response.getResponseCode() !== 200) return null;
  const fields = {};
  (response.getContentText().match(/<tr[\s\S]*?<\/tr>/gi) || []).forEach(row => {
    const cells = row.match(/<t[hd][^>]*>[\s\S]*?<\/t[hd]>/gi) || [];
    if (cells.length >= 2) fields[htmlToText(cells[0]).replace(/:$/, '').toLowerCase()] = htmlToText(cells[1]);
  });
  if (!fields['genotype']) return null;
  return { name: fields['name'] || name, genotype: fields['genotype'], url: strainOrderUrl('CGC', name) };
}

function totalFormula(r) {
  const q = colLetter(COL['Quantity']) + r, p = colLetter(COL['Unit price']) + r;
  return '=IF(OR(' + q + '="",' + p + '=""),"",' + q + '*' + p + ')';
}
function colLetter(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

function orderSummary(o) {
  const qty = toNumber(o['Quantity']) || 1;
  const price = o['Unit price'] === '' ? '' : toNumber(o['Unit price']);
  return [
    o['Type'] ? 'Type: ' + o['Type'] : '',
    'Item: ' + o['Item'],
    'Vendor: ' + (o['Vendor'] || '-') + (o['Catalog #'] ? '   Catalog #: ' + o['Catalog #'] : ''),
    'Quantity: ' + qty + (price !== '' ? '   Unit price: ' + money(price) + '   Total: ' + money(qty * price) : ''),
    'Grant: ' + (o['Grant'] || '-'),
    o['Link'] ? 'Link: ' + o['Link'] : '',
    o['Needed by'] ? 'Needed by: ' + formatDay(toDate(o['Needed by'])) : '',
    o['Notes'] ? 'Notes: ' + o['Notes'] : '',
  ].filter(l => l).join('\n');
}

function sheetLink() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getUrl() + '#gid=' + getOrders().getSheetId();
}

// Routes a new request to the people listed under "People who place orders" in Settings
function emailOrderers(o) {
  const to = getOrderers().map(p => p.email);
  if (!to.length) return 'No email sent: nobody is set (menu: Set who gets new-order emails).';
  try {
    MailApp.sendEmail({
      to: to.join(','),
      replyTo: o['Requester email'],
      subject: 'Order request ' + o['Order ID'] + ': ' + o['Item'] + ' (' + o['Requested by'] + ')',
      body: o['Requested by'] + ' asked for an order.\n\n' + orderSummary(o) +
        '\n\nWhen you have placed it, select the row and use the menu ' + CONFIG.menuName +
        ' > Mark selected orders as ordered (or set Status to "Ordered").\n\nOrders sheet: ' + sheetLink(),
    });
    return 'Sent to ' + to.join(', ') + '.';
  } catch (e) {
    console.log('Order email failed: ' + e);
    return 'The request was saved, but the email failed: ' + e.message;
  }
}

////////////////////////////////////////////////////////////////////////////
//    2. The orderer places it: Status -> Ordered, email the requester    //
////////////////////////////////////////////////////////////////////////////
function markOrdered() { markSelected('Ordered'); }
function markReceived() { markSelected('Received'); }
function markCancelled() { markSelected('Cancelled'); }

function selectedOrderRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== CONFIG.ordersSheet) throw new Error('Go to the ' + CONFIG.ordersSheet + ' tab and select the order rows first.');
  const rows = [];
  const ranges = sheet.getActiveRangeList() ? sheet.getActiveRangeList().getRanges() : [sheet.getActiveRange()];
  ranges.forEach(r => {
    for (let i = r.getRow(); i < r.getRow() + r.getNumRows(); i++) {
      if (i > CONFIG.headerRow && rows.indexOf(i) === -1) rows.push(i);
    }
  });
  return rows.filter(i => String(sheet.getRange(i, COL['Order ID']).getValue()).trim() !== '');
}

function markSelected(status) {
  const ui = SpreadsheetApp.getUi();
  const rows = selectedOrderRows();
  if (!rows.length) { ui.alert('Select one or more order rows (click anywhere in the row) first.'); return; }

  let confirmation = '';
  if (status === 'Ordered') {
    const res = ui.prompt('Mark ' + rows.length + ' order(s) as ordered',
      'Confirmation or PO number (optional). The person who asked gets an email.', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    confirmation = res.getResponseText().trim();
  } else if (status === 'Cancelled') {
    if (ui.alert('Cancel ' + rows.length + ' order(s)? The person who asked gets an email.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  }
  const email = Session.getActiveUser().getEmail() || '';
  const person = getOrderers().find(o => o.email.toLowerCase() === email.toLowerCase());
  const me = person && person.name ? person.name : email;
  const notes = rows.map(r => applyStatus(r, status, { confirmation: confirmation, by: me }));
  refreshAnalysisQuietly();
  SpreadsheetApp.getActive().toast(notes.join('\n'), 'Orders updated', 8);
}

// Sets the status of one row, fills in the dates and emails the requester.
// Used by the menu items and by the Status dropdown (see onOrderEdit).
function applyStatus(r, status, opts) {
  opts = opts || {};
  const sheet = getOrders();
  const today = new Date();
  withLock(() => {
    sheet.getRange(r, COL['Status']).setValue(status);
    if (status === 'Ordered') {
      if (!sheet.getRange(r, COL['Order date']).getValue()) sheet.getRange(r, COL['Order date']).setValue(today).setNumberFormat(DATE_FORMAT);
      if (opts.by && !sheet.getRange(r, COL['Ordered by']).getValue()) sheet.getRange(r, COL['Ordered by']).setValue(opts.by);
      if (opts.confirmation) sheet.getRange(r, COL['Confirmation / PO #']).setValue(opts.confirmation);
    }
    if (status === 'Received') {
      if (!sheet.getRange(r, COL['Order date']).getValue()) sheet.getRange(r, COL['Order date']).setValue(today).setNumberFormat(DATE_FORMAT);
      sheet.getRange(r, COL['Received date']).setValue(today).setNumberFormat(DATE_FORMAT);
    }
  });
  const values = sheet.getRange(r, 1, 1, COLS.length).getValues()[0];
  const o = {};
  COLS.forEach((c, i) => { o[c] = values[i]; });
  return o['Order ID'] + ': ' + emailRequester(o, status);
}

// Routes the update back to the person who asked
function emailRequester(o, status) {
  const to = String(o['Requester email']).trim();
  if (!isEmail(to)) return status + ' (no requester email on the row, nobody told)';
  const lines = {
    Ordered: 'Your order has been placed' + (o['Ordered by'] ? ' by ' + o['Ordered by'] : '') + '.' +
      (o['Confirmation / PO #'] ? '\nConfirmation / PO #: ' + o['Confirmation / PO #'] : ''),
    Received: 'Your order has arrived.' + (/^Strain /.test(String(o['Linked to'])) ?
      '\nIt is a strain: add it to the strain collection with "Add an existing strain (CGC / other labs)".' : ''),
    Cancelled: 'Your order request was cancelled. Reply to this email if you have questions.',
    Requested: 'Your order is back to "Requested" (not placed yet).',
  };
  try {
    MailApp.sendEmail({
      to: to,
      subject: o['Order ID'] + ' ' + status.toLowerCase() + ': ' + o['Item'],
      body: 'Hi ' + (o['Requested by'] || '') + ',\n\n' + (lines[status] || 'Status: ' + status) +
        '\n\n' + orderSummary(o) + '\n\nOrders sheet: ' + sheetLink(),
    });
    return status + ', ' + to + ' was emailed';
  } catch (e) {
    console.log('Requester email failed: ' + e);
    return status + ', but the email failed: ' + e.message;
  }
}

// Installable trigger (set up by "Set up the order sheets"): picking a status in the
// Status dropdown does the same as the menu items. Simple onEdit can't send email, so this
// one is installed and runs as the person who ran the setup.
function onOrderEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.ordersSheet) return;
  if (e.range.getColumn() !== COL['Status'] || e.range.getNumColumns() !== 1) return;
  const email = (e.user && e.user.getEmail && e.user.getEmail()) || '';
  const person = getOrderers().find(o => o.email.toLowerCase() === email.toLowerCase());
  const by = person && person.name ? person.name : email;
  for (let r = e.range.getRow(); r < e.range.getRow() + e.range.getNumRows(); r++) {
    if (r <= CONFIG.headerRow) continue;
    const status = String(sheet.getRange(r, COL['Status']).getValue()).trim();
    if (STATUSES.indexOf(status) === -1 || status === 'Requested') continue;
    if (e.range.getNumRows() === 1 && e.oldValue === status) continue;
    const note = applyStatus(r, status, { by: by });
    e.source.toast(note, 'Order updated', 6);
  }
  refreshAnalysisQuietly();
}

function installEditTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onOrderEdit');
  if (!exists) ScriptApp.newTrigger('onOrderEdit').forSpreadsheet(ss).onEdit().create();
}

////////////////////////////////////////////////////////////////////////////
//          3. Analysis tab: spend per grant / total, order frequency     //
////////////////////////////////////////////////////////////////////////////
// Spend counts orders that were placed (Ordered or Received). Requested ones are shown
// separately as "waiting", cancelled ones are left out.
function computeAnalysis(orders, grants, now) {
  now = now || new Date();
  const placed = orders.filter(o => o['Status'] === 'Ordered' || o['Status'] === 'Received');
  const waiting = orders.filter(o => o['Status'] === 'Requested' || String(o['Status']).trim() === '');
  const cost = o => (o['Unit price'] === '' ? 0 : (toNumber(o['Quantity']) || 1) * toNumber(o['Unit price']));
  const total = placed.reduce((s, o) => s + cost(o), 0);

  // Spend per grant
  const byGrant = {};
  grants.forEach(g => { byGrant[g.name] = { budget: g.budget, spent: 0, n: 0, waiting: 0 }; });
  const grantOf = o => String(o['Grant']).trim() || '(no grant)';
  placed.forEach(o => {
    const g = byGrant[grantOf(o)] = byGrant[grantOf(o)] || { budget: 0, spent: 0, n: 0, waiting: 0 };
    g.spent += cost(o); g.n++;
  });
  waiting.forEach(o => {
    const g = byGrant[grantOf(o)] = byGrant[grantOf(o)] || { budget: 0, spent: 0, n: 0, waiting: 0 };
    g.waiting += cost(o);
  });
  const grantRows = Object.keys(byGrant).map(name => {
    const g = byGrant[name];
    return [name, g.budget || '', g.spent, total ? g.spent / total : 0, g.budget ? g.budget - g.spent : '', g.waiting, g.n];
  }).sort((a, b) => b[2] - a[2]);

  // Orders per month (by request date), most recent first
  const monthKey = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  const months = [];
  for (let i = 0; i < CONFIG.monthsInAnalysis; i++) {
    months.push(monthKey(new Date(now.getFullYear(), now.getMonth() - i, 1)));
  }
  const monthRows = months.map(m => {
    const inMonth = orders.filter(o => o['Status'] !== 'Cancelled' && toDate(o['Request date']) && monthKey(toDate(o['Request date'])) === m);
    const placedIn = inMonth.filter(o => placed.indexOf(o) !== -1);
    return [m, inMonth.length, placedIn.reduce((s, o) => s + cost(o), 0)];
  });

  // Per person and per vendor
  const groupBy = (key) => {
    const map = {};
    orders.filter(o => o['Status'] !== 'Cancelled').forEach(o => {
      const k = String(o[key]).trim() || '(blank)';
      map[k] = map[k] || [k, 0, 0];
      map[k][1]++;
      if (placed.indexOf(o) !== -1) map[k][2] += cost(o);
    });
    return Object.keys(map).map(k => map[k]).sort((a, b) => b[1] - a[1] || b[2] - a[2]);
  };

  // Most-ordered items: how often they are ordered (average days between requests)
  const items = {};
  orders.filter(o => o['Status'] !== 'Cancelled').forEach(o => {
    const k = (String(o['Item']).trim() + (o['Catalog #'] ? ' (' + String(o['Catalog #']).trim() + ')' : '')).trim();
    if (!k) return;
    items[k] = items[k] || [];
    const d = toDate(o['Request date']);
    if (d) items[k].push(d);
  });
  const itemRows = Object.keys(items).map(k => {
    const ds = items[k].sort((a, b) => a - b);
    const gap = ds.length > 1 ? Math.round((ds[ds.length - 1] - ds[0]) / 86400000 / (ds.length - 1)) : '';
    return [k, ds.length, ds.length ? ds[ds.length - 1] : '', gap];
  }).sort((a, b) => b[1] - a[1]).slice(0, 25);

  // Days from request to order
  const lead = placed.map(o => {
    const a = toDate(o['Request date']), b = toDate(o['Order date']);
    return a && b ? (b - a) / 86400000 : null;
  }).filter(x => x !== null && x >= 0);
  const yearStart = new Date(now.getFullYear(), 0, 1);

  return {
    summary: [
      ['Total spend (placed orders)', total],
      ['Spend this year', placed.filter(o => (toDate(o['Order date']) || toDate(o['Request date']) || 0) >= yearStart).reduce((s, o) => s + cost(o), 0)],
      ['Orders placed', placed.length],
      ['Waiting to be ordered', waiting.length],
      ['Average days from request to order', lead.length ? Math.round(lead.reduce((s, x) => s + x, 0) / lead.length * 10) / 10 : ''],
    ],
    grantRows: grantRows,
    monthRows: monthRows,
    personRows: groupBy('Requested by'),
    vendorRows: groupBy('Vendor'),
    typeRows: groupBy('Type'),
    itemRows: itemRows,
  };
}

function refreshAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.analysisSheet);
  if (!sheet) sheet = ss.insertSheet(CONFIG.analysisSheet);
  const a = computeAnalysis(readOrders(), getGrants());

  sheet.getCharts().forEach(c => sheet.removeChart(c));
  sheet.getBandings().forEach(b => b.remove());
  sheet.clear();
  sheet.setHiddenGridlines(true);
  const moneyFmt = CONFIG.currency + '#,##0.00';

  sheet.getRange('A1').setValue(CONFIG.labName + ' lab orders: analysis')
    .setFontSize(14).setFontWeight('bold').setFontColor(THEME.header);
  sheet.getRange('A2').setValue('Updated ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') +
    '. Spend counts Ordered and Received orders. Updates on its own; after editing rows by hand, use the menu "Refresh analysis".')
    .setFontColor('#666666').setFontSize(9);

  let row = 4;
  // Each block: title, column names, rows. Returns the first data row and the row after the block.
  const block = (title, headers, rows, formats) => {
    sheet.getRange(row, 1).setValue(title).setFontWeight('bold').setFontSize(11).setFontColor(THEME.settingsText);
    const h = row + 1;
    sheet.getRange(h, 1, 1, headers.length).setValues([headers])
      .setBackground(THEME.header).setFontColor(THEME.headerText).setFontWeight('bold').setHorizontalAlignment('center');
    const data = rows.length ? rows : [[ 'Nothing yet' ].concat(headers.slice(1).map(() => ''))];
    const body = sheet.getRange(h + 1, 1, data.length, headers.length);
    // Formats first, so a month like "2026-10" stays text instead of turning into a date
    (formats || []).forEach((f, i) => { if (f) sheet.getRange(h + 1, i + 1, data.length, 1).setNumberFormat(f); });
    body.setValues(data).setBorder(null, null, true, null, null, true, THEME.grid, SpreadsheetApp.BorderStyle.SOLID);
    for (let i = 0; i < data.length; i += 2) sheet.getRange(h + 1 + i, 1, 1, headers.length).setBackground(THEME.stripe);
    const start = row;
    row = h + 1 + data.length + 2;
    return { header: h, first: h + 1, n: rows.length, start: start };
  };

  const sum = block('Summary', ['', 'Value'], a.summary, ['', '']);
  sheet.getRange(sum.first, 2, 2, 1).setNumberFormat(moneyFmt);
  sheet.getRange(sum.first, 2, a.summary.length, 1).setFontWeight('bold').setHorizontalAlignment('right');

  const g = block('Spend per grant',
    ['Grant', 'Budget', 'Spent', '% of total spend', 'Budget left', 'Waiting (not ordered yet)', 'Orders'],
    a.grantRows, ['', moneyFmt, moneyFmt, '0.0%', moneyFmt, moneyFmt, '0']);
  if (a.grantRows.length) {
    sheet.getRange(g.first + a.grantRows.length, 1, 1, 7)
      .setValues([['Total', '', a.summary[0][1], a.summary[0][1] ? 1 : 0, '', a.grantRows.reduce((s, r) => s + r[5], 0), a.grantRows.reduce((s, r) => s + r[6], 0)]])
      .setFontWeight('bold').setBorder(true, null, null, null, null, null, THEME.header, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(g.first + a.grantRows.length, 2, 1, 6).setNumberFormats([[moneyFmt, moneyFmt, '0.0%', moneyFmt, moneyFmt, '0']]);
    row++;
  }
  const m = block('Orders per month (last ' + CONFIG.monthsInAnalysis + ' months)', ['Month', 'Orders', 'Spend'], a.monthRows, ['@', '0', moneyFmt]);
  block('Orders per type (strain, plasmid, oligo, supplies)', ['Type', 'Orders', 'Spend'], a.typeRows, ['', '0', moneyFmt]);
  const p = block('Orders per person', ['Person', 'Orders', 'Spend'], a.personRows, ['', '0', moneyFmt]);
  block('Orders per vendor', ['Vendor', 'Orders', 'Spend'], a.vendorRows, ['', '0', moneyFmt]);
  block('Most-ordered items', ['Item', 'Times ordered', 'Last requested', 'Ordered about every (days)'], a.itemRows, ['', '0', DATE_FORMAT, '0']);

  [260, 110, 110, 120, 110, 150, 70].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange(1, 1, row, 7).setFontFamily('Arial');

  // Charts to the right of the tables
  if (a.grantRows.some(r => r[2] > 0)) {
    sheet.insertChart(sheet.newChart().asPieChart()
      .addRange(sheet.getRange(g.header, 1, a.grantRows.length + 1, 1))
      .addRange(sheet.getRange(g.header, 3, a.grantRows.length + 1, 1))
      .setNumHeaders(1).setOption('title', 'Spend per grant').setOption('pieHole', 0.4)
      .setPosition(4, 9, 0, 0).setOption('width', 480).setOption('height', 300).build());
  }
  sheet.insertChart(sheet.newChart().asColumnChart()
    .addRange(sheet.getRange(m.header, 1, a.monthRows.length + 1, 2))
    .setNumHeaders(1).setOption('title', 'Orders per month').setOption('legend', { position: 'none' })
    .setOption('hAxis', { direction: -1 }).setOption('colors', [THEME.header])
    .setPosition(21, 9, 0, 0).setOption('width', 480).setOption('height', 300).build());
  return a;
}

// Refresh after a change without ever blocking the change itself
function refreshAnalysisQuietly() {
  try { refreshAnalysis(); } catch (e) { console.log('Analysis refresh failed: ' + e); }
}

function refreshAnalysisFromMenu() {
  refreshAnalysis();
  SpreadsheetApp.getActive().toast('Analysis updated.');
}

////////////////////////////////////////////////////////////////////////////
//                   One-time setup and formatting                        //
////////////////////////////////////////////////////////////////////////////
function setupOrderSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let orders = ss.getSheetByName(CONFIG.ordersSheet);
  if (!orders) orders = ss.insertSheet(CONFIG.ordersSheet, 0);
  orders.getRange(1, 1).setValue(CONFIG.labName + ' lab orders');
  orders.getRange(CONFIG.headerRow, 1, 1, COLS.length).setValues([COLS]);

  let settings = ss.getSheetByName(CONFIG.settingsSheet);
  if (!settings) {
    settings = ss.insertSheet(CONFIG.settingsSheet);
    settings.getRange('A3:C4').setValues([['Example grant R01 123456', 50000, 'Replace with your grants'], ['Startup funds', 20000, '']]);
  }
  settings.getRange('A1').setValue('Grants: the Grant dropdown on the Orders tab comes from this list');
  settings.getRange('E1').setValue('People who place orders: they get an email for every new request');
  settings.getRange('A2:C2').setValues([['Grant', 'Budget', 'Notes']]);
  settings.getRange('E2:F2').setValues([['Name', 'Email']]);

  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  formatSheets();
  refreshAnalysis();
  installEditTrigger();
  ss.setActiveSheet(settings);
  SpreadsheetApp.getUi().alert('The order sheets are ready.\n\n1. In Settings, list your grants (with budgets) and the people who place orders.\n' +
    '2. Use the menu ' + CONFIG.menuName + ' > Request an order.');
}

function formatSheets() {
  formatOrdersSheet();
  formatSettingsSheet();
  SpreadsheetApp.getActive().toast('Sheets formatted.');
}

function formatOrderRow(sheet, r) {
  [COL['Request date'], COL['Needed by'], COL['Order date'], COL['Received date']]
    .forEach(c => sheet.getRange(r, c).setNumberFormat(DATE_FORMAT));
  [COL['Unit price'], COL['Total']].forEach(c => sheet.getRange(r, c).setNumberFormat(CONFIG.currency + '#,##0.00'));
}

function formatOrdersSheet() {
  const s = getOrders();
  const h = CONFIG.headerRow;
  const n = COLS.length;
  // Sheets made before the Type column existed: make room for it so "Linked to" keeps its data
  if (String(s.getRange(h, COL['Type']).getValue()).trim() === 'Linked to') s.insertColumnBefore(COL['Type']);
  s.getRange(h, 1, 1, n).setValues([COLS]);
  s.getRange(1, 1, 1, n).setBackground(THEME.settings);
  s.getRange(1, 1).setFontSize(13).setFontWeight('bold').setFontColor(THEME.settingsText);
  s.setRowHeight(1, 30);

  const rows = Math.max(s.getMaxRows() - h, 1);
  const table = s.getRange(h, 1, rows + 1, n);
  s.getBandings().forEach(b => b.remove());
  table.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle').setWrap(true);
  table.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(THEME.header).setFirstRowColor('#ffffff').setSecondRowColor(THEME.stripe);
  s.getRange(h, 1, 1, n).setFontColor(THEME.headerText).setFontWeight('bold').setHorizontalAlignment('center');
  s.setRowHeight(h, 36);
  s.getRange(h + 1, 1, rows, n).setHorizontalAlignment('center')
    .setBorder(null, null, true, null, null, true, THEME.grid, SpreadsheetApp.BorderStyle.SOLID);
  ['Item', 'Notes', 'Requester email', 'Link'].forEach(c => s.getRange(h + 1, COL[c], rows, 1).setHorizontalAlignment('left'));
  s.getRange(h + 1, COL['Order ID'], rows, 1).setFontWeight('bold');
  [COL['Request date'], COL['Needed by'], COL['Order date'], COL['Received date']]
    .forEach(c => s.getRange(h + 1, c, rows, 1).setNumberFormat(DATE_FORMAT));
  [COL['Unit price'], COL['Total']].forEach(c => s.getRange(h + 1, c, rows, 1).setNumberFormat(CONFIG.currency + '#,##0.00'));
  s.getRange(h + 1, COL['Link'], rows, 1).setWrap(false);

  // Dropdowns: Status, and Grant from the Settings tab
  const status = s.getRange(h + 1, COL['Status'], rows, 1);
  status.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  s.getRange(h + 1, COL['Type'], rows, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ORDER_TYPES, true).build());
  const settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.settingsSheet);
  if (settings) {
    s.getRange(h + 1, COL['Grant'], rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(settings.getRange('A3:A'), true).setAllowInvalid(true).build());
  }
  const a1 = status.getA1Notation();
  const rules = s.getConditionalFormatRules().filter(rule => !rule.getRanges().some(g => g.getA1Notation() === a1));
  STATUSES.forEach(st => rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(st)
    .setBackground(STATUS_COLORS[st][0]).setFontColor(STATUS_COLORS[st][1]).setBold(true).setRanges([status]).build()));
  s.setConditionalFormatRules(rules);

  if (!s.getFilter()) table.createFilter();
  s.setFrozenRows(h);
  s.setFrozenColumns(1);
  COLUMN_WIDTHS.forEach((w, i) => s.setColumnWidth(i + 1, w));
}

function formatSettingsSheet() {
  const s = getSettings();
  s.getRange('A1:F1').setFontWeight('bold').setFontColor(THEME.settingsText).setWrap(true);
  s.setRowHeight(1, 36);
  ['A2:C2', 'E2:F2'].forEach(a1 => s.getRange(a1).setBackground(THEME.header).setFontColor(THEME.headerText)
    .setFontWeight('bold').setHorizontalAlignment('center'));
  s.getRange('B3:B').setNumberFormat(CONFIG.currency + '#,##0');
  [260, 110, 220, 30, 160, 240].forEach((w, i) => s.setColumnWidth(i + 1, w));
  s.setFrozenRows(2);
}

////////////////////////////////////////////////////////////////////////////
//                                  Menu                                  //
////////////////////////////////////////////////////////////////////////////
function onOpen() {
  SpreadsheetApp.getUi().createMenu(CONFIG.menuName)
    .addItem('Request an order', 'startOrderForm')
    .addSeparator()
    .addItem('Mark selected orders as ordered', 'markOrdered')
    .addItem('Mark selected orders as received', 'markReceived')
    .addItem('Cancel selected orders', 'markCancelled')
    .addSeparator()
    .addItem('Set who gets new-order emails', 'setOrderEmails')
    .addSeparator()
    .addItem('Refresh analysis', 'refreshAnalysisFromMenu')
    .addSeparator()
    .addItem('Set up the order sheets', 'setupOrderSheets')
    .addItem('Format the sheets', 'formatSheets')
    .addToUi();
}
