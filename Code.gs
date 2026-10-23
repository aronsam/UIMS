/**
 * UIMS — Universal Inventory Management System (integrated MVP, v1.1 corrected)
 * Google Apps Script + Google Sheets. Basahin ang README bago mag-deploy.
 * Gamitin lang sa dummy/test data hanggang hindi pa nata-test nang buo.
 */
const U = {
  DB: 'UIMS_DATABASE_ID',
  TZ: 'Asia/Manila',
  SESSION_SECONDS: 21600,   // 6 hours (max ng CacheService)
  MAX_FAILS: 5,             // failed logins bago mag-lock
  LOCK_SECONDS: 600,        // 10 minutes
  // Mga column na HINDI gagawing plain-text (numbers/booleans). Lahat ng iba ay text.
  NON_TEXT: ['TrackExpiry', 'MustChangePassword', 'Allowed', 'ReorderLevel', 'ConversionToBase',
    'Qty', 'QtyBase', 'UnitCost', 'QtyIn', 'QtyRemaining', 'ExpectedQty', 'CountedQty', 'VarianceQty'],
  SHEETS: {
    Businesses: ['BusinessID', 'BusinessName', 'Status', 'CreatedAt'],
    Users: ['UserID', 'BusinessID', 'Username', 'FullName', 'PasswordHash', 'Salt', 'Role', 'Status', 'MustChangePassword', 'CreatedAt', 'LastLogin'],
    Permissions: ['PermissionKey', 'Description'],
    UserPermissions: ['BusinessID', 'UserID', 'PermissionKey', 'Allowed'],
    Locations: ['LocationID', 'BusinessID', 'Name', 'Status', 'CreatedAt'],
    Categories: ['CategoryID', 'BusinessID', 'Name', 'CreatedAt'],
    Units: ['UnitID', 'BusinessID', 'Name', 'Abbreviation', 'CreatedAt'],
    Items: ['ItemID', 'BusinessID', 'SKU', 'Name', 'CategoryID', 'BaseUnitID', 'ReorderLevel', 'TrackExpiry', 'Status', 'CreatedAt'],
    ItemUnits: ['BusinessID', 'ItemID', 'UnitID', 'ConversionToBase'],
    Suppliers: ['SupplierID', 'BusinessID', 'Name', 'Contact', 'Status', 'CreatedAt'],
    Transactions: ['TransactionID', 'BusinessID', 'Type', 'Reference', 'LocationID', 'FromLocationID', 'ToLocationID', 'Status', 'CreatedBy', 'CreatedAt', 'Notes', 'ReversalOf'],
    TransactionLines: ['LineID', 'BusinessID', 'TransactionID', 'ItemID', 'Qty', 'UnitID', 'QtyBase', 'UnitCost', 'BatchNo', 'ExpiryDate', 'Notes'],
    StockBalances: ['BusinessID', 'LocationID', 'ItemID', 'QtyBase', 'UpdatedAt'],
    FIFOLayers: ['LayerID', 'BusinessID', 'LocationID', 'ItemID', 'TransactionID', 'ReceivedAt', 'QtyIn', 'QtyRemaining', 'UnitCost', 'BatchNo', 'ExpiryDate'],
    Approvals: ['ApprovalID', 'BusinessID', 'TransactionID', 'RequestedBy', 'Approver', 'Status', 'Reason', 'CreatedAt', 'UpdatedAt'],
    Transfers: ['TransferID', 'BusinessID', 'TransactionID', 'FromLocationID', 'ToLocationID', 'Status', 'DispatchedAt', 'ReceivedAt', 'DiscrepancyNotes'],
    StockCounts: ['CountID', 'BusinessID', 'LocationID', 'Status', 'CreatedBy', 'CreatedAt', 'PostedAt', 'Notes'],
    StockCountLines: ['LineID', 'BusinessID', 'CountID', 'ItemID', 'ExpectedQty', 'CountedQty', 'VarianceQty'],
    AuditLog: ['AuditID', 'BusinessID', 'UserID', 'Action', 'Entity', 'EntityID', 'Details', 'CreatedAt'],
    SyncLog: ['SyncID', 'BusinessID', 'UserID', 'Status', 'Message', 'CreatedAt'],
    Settings: ['BusinessID', 'Key', 'Value']
  }
};

const PERMS = [
  ['manage_users', 'Manage users and roles'], ['manage_items', 'Manage items and master data'],
  ['manage_locations', 'Manage locations'], ['receive_stock', 'Receive stock'],
  ['issue_stock', 'Issue stock'], ['adjust_stock', 'Adjust stock'],
  ['approve_transactions', 'Approve transactions'], ['view_cost', 'View inventory cost'],
  ['view_reports', 'View reports'], ['transfer_stock', 'Request and dispatch transfers'],
  ['confirm_transfers', 'Confirm transfer receipts'], ['stock_count', 'Create and post stock counts'],
  ['manage_settings', 'Manage business settings']
];

// Sino ang puwedeng MAGBASA ng master lists (para gumana ang dropdowns sa forms)
const OPS_PERMS = ['receive_stock', 'issue_stock', 'adjust_stock', 'transfer_stock', 'confirm_transfers', 'stock_count', 'view_reports'];
const MASTER = {
  items: { sheet: 'Items', write: 'manage_items', read: ['manage_items'].concat(OPS_PERMS) },
  locations: { sheet: 'Locations', write: 'manage_locations', read: ['manage_locations', 'manage_items'].concat(OPS_PERMS) },
  categories: { sheet: 'Categories', write: 'manage_items', read: ['manage_items'].concat(OPS_PERMS) },
  units: { sheet: 'Units', write: 'manage_items', read: ['manage_items'].concat(OPS_PERMS) },
  suppliers: { sheet: 'Suppliers', write: 'manage_items', read: ['manage_items', 'receive_stock'] }
};

/* ============================ WEB APP ============================ */

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('UIMS — Universal Inventory Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

/* ============================ SETUP (OWNER ONLY) ============================ */

let SS_ = null; // cache ng spreadsheet sa loob ng iisang execution
function db_() {
  if (SS_) return SS_;
  const id = PropertiesService.getScriptProperties().getProperty(U.DB);
  if (!id) throw new Error('Database not initialized. Run setupUIMSDatabase() in Apps Script first.');
  SS_ = SpreadsheetApp.openById(id);
  return SS_;
}

/** Hindi ito matatawag ng ibang web user: kailangang ang owner mismo ang tumatakbo. */
function assertOwner_() {
  const owner = Session.getEffectiveUser().getEmail();
  const current = Session.getActiveUser().getEmail();
  if (!owner || !current || owner !== current) throw new Error('Owner lang ang puwedeng gumamit ng function na ito.');
}

/** Patakbuhin nang manu-mano mula sa editor. Safe i-rerun (hindi gagawa ng bagong database). */
function setupUIMSDatabase() {
  assertOwner_();
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty(U.DB), ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('UIMS_Database');
    ss.setSpreadsheetTimeZone(U.TZ);
    id = ss.getId();
    p.setProperty(U.DB, id);
  }
  const first = ss.getSheets()[0];
  if (first.getName() === 'Sheet1' && !ss.getSheetByName('Businesses')) first.setName('Businesses');

  Object.keys(U.SHEETS).forEach(name => {
    const headers = U.SHEETS[name];
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#163B2C').setFontColor('#FFFFFF');
    }
    // Plain text ang mga text columns: hindi masisira ang SKU/barcode (leading zeros), petsa, hash, atbp.
    headers.forEach((h, i) => {
      if (U.NON_TEXT.indexOf(h) === -1) sh.getRange(2, i + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
    });
    sh.setFrozenRows(1);
  });

  const ps = ss.getSheetByName('Permissions');
  if (ps.getLastRow() === 1) ps.getRange(2, 1, PERMS.length, 2).setValues(PERMS);
  const set = ss.getSheetByName('Settings');
  if (set.getLastRow() === 1) {
    set.getRange(2, 1, 4, 3).setValues([['*', 'timezone', U.TZ], ['*', 'currency', 'PHP'], ['*', 'costing', 'FIFO'], ['*', 'allow_negative_stock', 'TRUE']]);
  }

  let msg = 'Database ready. Check Execution log for the database URL.';
  if (!p.getProperty('UIMS_ADMIN_CREATED')) {
    let token = p.getProperty('UIMS_BOOTSTRAP_TOKEN');
    if (!token) {
      token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
      p.setProperty('UIMS_BOOTSTRAP_TOKEN', token);
    }
    Logger.log('ONE-TIME INITIAL ADMIN SETUP TOKEN: ' + token);
    msg += ' Setup token is in the log.';
  }
  Logger.log('UIMS database URL: ' + ss.getUrl());
  return msg;
}

/** Kung nawala ang token sa log, patakbuhin ito (hangga't wala pang admin). */
function showSetupToken() {
  assertOwner_();
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('UIMS_ADMIN_CREATED')) { Logger.log('Admin exists na. Walang token.'); return 'Admin already created.'; }
  const t = p.getProperty('UIMS_BOOTSTRAP_TOKEN');
  if (!t) throw new Error('Run setupUIMSDatabase() first.');
  Logger.log('ONE-TIME INITIAL ADMIN SETUP TOKEN: ' + t);
  return 'Token logged.';
}

function getDatabaseUrlForOwner() {
  assertOwner_();
  const url = db_().getUrl();
  Logger.log(url);
  return url;
}

function getSetupStatus() {
  const p = PropertiesService.getScriptProperties();
  return { initialized: !!p.getProperty(U.DB), adminExists: !!p.getProperty('UIMS_ADMIN_CREATED') };
}

function bootstrapFirstAdmin(data) {
  const p = PropertiesService.getScriptProperties();
  data = data || {};
  const business = String(data.businessName || '').trim(), username = clean_(data.username),
    fullName = String(data.fullName || '').trim(), password = String(data.password || '');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (p.getProperty('UIMS_ADMIN_CREATED')) throw new Error('Initial admin already exists.');
    const expected = p.getProperty('UIMS_BOOTSTRAP_TOKEN');
    if (!expected || String(data.token || '') !== expected) throw new Error('Invalid one-time setup token.');
    if (!business || !username || !fullName) throw new Error('Complete all fields.');
    validatePassword_(password);
    const ss = db_(), bid = id_('BIZ'), uid = id_('USR'), salt = Utilities.getUuid();
    ss.getSheetByName('Businesses').appendRow([bid, business, 'Active', now_()]);
    ss.getSheetByName('Users').appendRow([uid, bid, username, fullName, hash_(password, salt), salt, 'BusinessAdmin', 'Active', true, now_(), '']);
    ss.getSheetByName('Locations').appendRow([id_('LOC'), bid, 'Main Warehouse', 'Active', now_()]);
    const set = ss.getSheetByName('Settings');
    set.appendRow([bid, 'currency', 'PHP']);
    set.appendRow([bid, 'costing', 'FIFO']);
    set.appendRow([bid, 'allow_negative_stock', 'TRUE']);
    p.setProperty('UIMS_ADMIN_CREATED', 'true');
    p.deleteProperty('UIMS_BOOTSTRAP_TOKEN');
    audit_(bid, uid, 'BOOTSTRAP', 'Business', bid, 'Initial business and admin created');
    return { ok: true, message: 'Business and admin created. Sign in with your new credentials; first login requires a password change.' };
  } finally { lock.releaseLock(); }
}

/* ============================ AUTH / SESSION ============================ */

function login(username, password) {
  username = clean_(username);
  password = String(password || '');
  if (!username || !password) throw new Error('Enter username and password.');
  const cache = CacheService.getScriptCache(), failKey = 'fail_' + username.toLowerCase();
  if (num_(cache.get(failKey)) >= U.MAX_FAILS) throw new Error('Too many failed attempts. Try again in 10 minutes.');
  const sh = db_().getSheetByName('Users'), d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    const r = d[i];
    if (String(r[2]).toLowerCase() !== username.toLowerCase()) continue;
    if (r[7] === 'Active' && hash_(password, String(r[5])) === String(r[4])) {
      cache.remove(failKey);
      const token = Utilities.getUuid() + Utilities.getUuid();
      const ses = { userId: r[0], businessId: r[1], username: r[2], fullName: r[3], role: r[6], mustChange: isTrue_(r[8]) };
      cache.put('sess_' + token, JSON.stringify(ses), U.SESSION_SECONDS);
      sh.getRange(i + 1, 11).setValue(now_());
      audit_(r[1], r[0], 'LOGIN', 'User', r[0], 'Successful login');
      return { token: token, user: { username: r[2], fullName: r[3], role: r[6], mustChange: ses.mustChange, perms: Array.from(permSet_(ses)) } };
    }
  }
  cache.put(failKey, String(num_(cache.get(failKey)) + 1), U.LOCK_SECONDS);
  throw new Error('Invalid username or password, or account disabled.');
}

function logout(token) {
  const s = session_(token);
  CacheService.getScriptCache().remove('sess_' + token);
  audit_(s.businessId, s.userId, 'LOGOUT', 'User', s.userId, 'Logout');
  return { ok: true };
}

function session_(token) {
  if (!token) throw new Error('Session expired. Please sign in again.');
  const raw = CacheService.getScriptCache().get('sess_' + token);
  if (!raw) throw new Error('Session expired. Please sign in again.');
  const s = JSON.parse(raw);
  const r = db_().getSheetByName('Users').getDataRange().getValues().slice(1).find(x => x[0] === s.userId && x[1] === s.businessId);
  if (!r || r[7] !== 'Active') throw new Error('Account is inactive.');
  s.role = r[6];                 // laging fresh mula sa sheet
  s.mustChange = isTrue_(r[8]);
  return s;
}

function permSet_(s) {
  if (s._perms) return s._perms;
  const set = new Set();
  if (s.role === 'BusinessAdmin') PERMS.forEach(p => set.add(p[0]));
  else db_().getSheetByName('UserPermissions').getDataRange().getValues().slice(1).forEach(r => {
    if (r[0] === s.businessId && r[1] === s.userId && isTrue_(r[3])) set.add(r[2]);
  });
  s._perms = set;
  return set;
}
function hasPerm_(s, p) { return permSet_(s).has(p); }
function hasAny_(s, list) { return list.some(p => hasPerm_(s, p)); }

function require_(token, perm, allowMustChange) {
  const s = session_(token);
  if (s.mustChange && !allowMustChange) throw new Error('Change your temporary password before continuing.');
  if (perm && !hasPerm_(s, perm)) throw new Error('Permission denied: ' + perm);
  return s;
}
function requireAny_(token, list, label) {
  const s = session_(token);
  if (s.mustChange) throw new Error('Change your temporary password before continuing.');
  if (!hasAny_(s, list)) throw new Error('Permission denied' + (label ? ' to view ' + label : '') + '.');
  return s;
}

function changePassword(token, oldPassword, newPassword) {
  const s = require_(token, null, true);
  validatePassword_(newPassword);
  const sh = db_().getSheetByName('Users'), d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] === s.userId) {
      if (hash_(String(oldPassword || ''), String(d[i][5])) !== String(d[i][4])) throw new Error('Current password is incorrect.');
      if (String(oldPassword) === String(newPassword)) throw new Error('New password must be different from the current one.');
      const salt = Utilities.getUuid();
      sh.getRange(i + 1, 5, 1, 2).setValues([[hash_(newPassword, salt), salt]]);
      sh.getRange(i + 1, 9).setValue(false);
      audit_(s.businessId, s.userId, 'CHANGE_PASSWORD', 'User', s.userId, 'Password changed');
      return { ok: true, message: 'Password changed.' };
    }
  }
  throw new Error('User not found.');
}

/* ============================ MASTER DATA ============================ */

function dashboard(token) {
  const s = require_(token, null, false), ss = db_(), bid = s.businessId;
  const items = rows_(ss, 'Items', bid).filter(i => i[8] === 'Active'), locs = rows_(ss, 'Locations', bid), bal = rows_(ss, 'StockBalances', bid);
  let negative = 0, low = 0;
  bal.forEach(b => { if (num_(b[3]) < 0) negative++; });
  items.forEach(it => {
    const total = bal.filter(b => b[2] === it[0]).reduce((a, b) => a + num_(b[3]), 0);
    if (total <= 0 || total <= num_(it[6])) low++;
  });
  return { items: items.length, locations: locs.length, negative: negative, low: low, user: s.fullName, role: s.role };
}

function listData(token, kind) {
  const m = MASTER[kind];
  if (!m) throw new Error('Unknown data type.');
  const s = requireAny_(token, m.read, kind);
  return rows_(db_(), m.sheet, s.businessId).map(r => ({ id: r[0], values: r }));
}

function listItemUnits(token) {
  const s = requireAny_(token, MASTER.items.read, 'item units');
  return rows_(db_(), 'ItemUnits', s.businessId).map(r => ({ itemId: r[1], unitId: r[2], factor: num_(r[3]) }));
}

function saveMaster(token, kind, data) {
  const m = MASTER[kind];
  if (!m) throw new Error('Unknown data type.');
  const s = require_(token, m.write, false), ss = db_(), bid = s.businessId, sh = ss.getSheetByName(m.sheet);
  data = data || {};
  const name = String(data.name || '').trim();
  if (!name) throw new Error('Name is required.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existing = rows_(ss, m.sheet, bid);
    const id = data.id || id_(kind.slice(0, 3).toUpperCase());
    if (data.id && !existing.some(r => r[0] === id)) throw new Error('Record not found.');
    // Duplicate check
    if (existing.some(r => r[0] !== id && String(r[2 + (kind === 'items' ? 1 : 0)]).toLowerCase() === name.toLowerCase()))
      throw new Error(name + ' already exists.');
    const old = existing.find(r => r[0] === id);
    const stamp = old ? old[old.length - 1] : now_();
    let row;
    if (kind === 'items') {
      const sku = String(data.sku || '').trim();
      if (sku && existing.some(r => r[0] !== id && String(r[2]).toLowerCase() === sku.toLowerCase())) throw new Error('SKU already exists.');
      const baseUnitId = String(data.baseUnitId || '');
      if (!rows_(ss, 'Units', bid).some(u => u[0] === baseUnitId)) throw new Error('Choose a valid Base Unit.');
      const categoryId = String(data.categoryId || '');
      if (categoryId && !rows_(ss, 'Categories', bid).some(c => c[0] === categoryId)) throw new Error('Invalid category.');
      row = [id, bid, sku, name, categoryId, baseUnitId, Math.max(0, num_(data.reorderLevel)), !!data.trackExpiry, old ? old[8] : 'Active', stamp];
    } else if (kind === 'locations') row = [id, bid, name, old ? old[3] : 'Active', stamp];
    else if (kind === 'categories') row = [id, bid, name, stamp];
    else if (kind === 'units') row = [id, bid, name, String(data.abbreviation || name).trim(), stamp];
    else row = [id, bid, name, String(data.contact || ''), old ? old[4] : 'Active', stamp];

    if (data.id) {
      const all = sh.getDataRange().getValues();
      const idx = all.findIndex((r, i) => i > 0 && r[0] === id && r[1] === bid);
      sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);
    } else sh.appendRow(row);
    audit_(bid, s.userId, 'SAVE_MASTER', m.sheet, id, JSON.stringify({ name: name }));
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}

/** Mag-set ng alternate unit para sa item (hal. 1 Box = 12 Piece). */
function saveItemUnit(token, d) {
  const s = require_(token, 'manage_items', false), ss = db_(), bid = s.businessId;
  d = d || {};
  const item = rows_(ss, 'Items', bid).find(r => r[0] === d.itemId);
  if (!item) throw new Error('Item not found.');
  if (!rows_(ss, 'Units', bid).some(u => u[0] === d.unitId)) throw new Error('Unit not found.');
  if (d.unitId === item[5]) throw new Error('That is already the base unit of the item.');
  const factor = Number(d.factor);
  if (!(factor > 0)) throw new Error('Conversion must be greater than zero.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = ss.getSheetByName('ItemUnits'), all = sh.getDataRange().getValues();
    const idx = all.findIndex((r, i) => i > 0 && r[0] === bid && r[1] === d.itemId && r[2] === d.unitId);
    if (idx > 0) sh.getRange(idx + 1, 4).setValue(factor); else sh.appendRow([bid, d.itemId, d.unitId, factor]);
    audit_(bid, s.userId, 'SAVE_ITEM_UNIT', 'Item', d.itemId, d.unitId + ' x' + factor);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ============================ USERS & ACCESS ============================ */

function getPermissions(token) {
  require_(token, 'manage_users', false);
  return PERMS.map(p => ({ key: p[0], description: p[1] }));
}

function listUsers(token) {
  const s = require_(token, 'manage_users', false);
  return rows_(db_(), 'Users', s.businessId).map(r => ({ id: r[0], username: r[2], fullName: r[3], role: r[6], status: r[7], mustChange: isTrue_(r[8]) }));
}

function getUserPermissions(token, userId) {
  const s = require_(token, 'manage_users', false);
  const u = rows_(db_(), 'Users', s.businessId).find(r => r[0] === userId);
  if (!u) throw new Error('User not found.');
  return Array.from(permSet_({ businessId: s.businessId, userId: userId, role: u[6] }));
}

function createUser(token, d) {
  const s = require_(token, 'manage_users', false), ss = db_();
  d = d || {};
  const username = clean_(d.username), fullName = String(d.fullName || '').trim(), password = String(d.password || '');
  if (!username || !fullName) throw new Error('Username and full name are required.');
  validatePassword_(password);
  if (!['WarehouseStaff', 'Supervisor', 'BusinessAdmin'].includes(d.role)) throw new Error('Invalid role.');
  if (d.role === 'BusinessAdmin' && s.role !== 'BusinessAdmin') throw new Error('Only a Business Admin can create another Business Admin.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = ss.getSheetByName('Users'), all = sh.getDataRange().getValues();
    // Global unique ang username dahil walang business selector sa login.
    if (all.slice(1).some(r => String(r[2]).toLowerCase() === username.toLowerCase())) throw new Error('Username already exists.');
    const uid = id_('USR'), salt = Utilities.getUuid();
    sh.appendRow([uid, s.businessId, username, fullName, hash_(password, salt), salt, d.role, 'Active', true, now_(), '']);
    writePerms_(ss, s.businessId, uid, Array.isArray(d.permissions) ? d.permissions : []);
    audit_(s.businessId, s.userId, 'CREATE_USER', 'User', uid, username);
    return { ok: true, userId: uid, message: 'User created. Share temporary credentials privately.' };
  } finally { lock.releaseLock(); }
}

function writePerms_(ss, bid, uid, list) {
  const rows = PERMS.map(p => [bid, uid, p[0], list.indexOf(p[0]) !== -1]);
  const sh = ss.getSheetByName('UserPermissions');
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
}

/** Non-admin na may manage_users ay hindi puwedeng galawin ang admin accounts o ang sarili niyang access. */
function guardUserChange_(s, target) {
  if (!target) throw new Error('User not found.');
  if (s.role !== 'BusinessAdmin' && (target[6] === 'BusinessAdmin' || target[0] === s.userId))
    throw new Error('Only a Business Admin can change this account.');
}

function setUserStatus(token, userId, status) {
  const s = require_(token, 'manage_users', false);
  if (!['Active', 'Disabled'].includes(status)) throw new Error('Invalid status.');
  const sh = db_().getSheetByName('Users'), d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) if (d[i][0] === userId && d[i][1] === s.businessId) {
    guardUserChange_(s, d[i]);
    if (userId === s.userId && status === 'Disabled') throw new Error('You cannot disable your own account.');
    sh.getRange(i + 1, 8).setValue(status);
    audit_(s.businessId, s.userId, 'USER_STATUS', 'User', userId, status);
    return { ok: true };
  }
  throw new Error('User not found.');
}

function setPermissions(token, userId, permissions) {
  const s = require_(token, 'manage_users', false), ss = db_();
  const target = rows_(ss, 'Users', s.businessId).find(r => r[0] === userId);
  guardUserChange_(s, target);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = ss.getSheetByName('UserPermissions'), all = sh.getDataRange().getValues();
    for (let i = all.length - 1; i >= 1; i--) if (all[i][0] === s.businessId && all[i][1] === userId) sh.deleteRow(i + 1);
    writePerms_(ss, s.businessId, userId, Array.isArray(permissions) ? permissions : []);
    audit_(s.businessId, s.userId, 'SET_PERMISSIONS', 'User', userId, JSON.stringify(permissions || []));
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/** Reset ng password ng user (temporary password; kailangan niyang palitan sa susunod na login). */
function resetUserPassword(token, userId, tempPassword) {
  const s = require_(token, 'manage_users', false);
  validatePassword_(String(tempPassword || ''));
  const sh = db_().getSheetByName('Users'), d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) if (d[i][0] === userId && d[i][1] === s.businessId) {
    guardUserChange_(s, d[i]);
    const salt = Utilities.getUuid();
    sh.getRange(i + 1, 5, 1, 2).setValues([[hash_(String(tempPassword), salt), salt]]);
    sh.getRange(i + 1, 9).setValue(true);
    CacheService.getScriptCache().remove('fail_' + String(d[i][2]).toLowerCase());
    audit_(s.businessId, s.userId, 'RESET_PASSWORD', 'User', userId, '');
    return { ok: true, message: 'Temporary password set. User must change it at next login.' };
  }
  throw new Error('User not found.');
}

/* ============================ INVENTORY ============================ */

function inventory(token, locationId) {
  const s = require_(token, 'view_reports', false), ss = db_(), bid = s.businessId;
  const items = rows_(ss, 'Items', bid), bal = rows_(ss, 'StockBalances', bid), layers = rows_(ss, 'FIFOLayers', bid),
    units = rows_(ss, 'Units', bid), cats = rows_(ss, 'Categories', bid);
  const canCost = hasPerm_(s, 'view_cost');
  return items.map(i => {
    const qty = r6_(bal.filter(b => b[2] === i[0] && (!locationId || b[1] === locationId)).reduce((a, b) => a + num_(b[3]), 0));
    const cost = layers.filter(l => l[3] === i[0] && (!locationId || l[2] === locationId) && num_(l[7]) > 0)
      .reduce((a, l) => a + num_(l[7]) * num_(l[8]), 0);
    const reorder = num_(i[6]), unit = units.find(u => u[0] === i[5]), cat = cats.find(c => c[0] === i[4]);
    return {
      itemId: i[0], sku: i[2], name: i[3], category: cat ? cat[2] : '', unit: unit ? unit[3] : '', reorderLevel: reorder, qty: qty,
      valuation: canCost ? r6_(cost) : null,
      status: qty < 0 ? 'NEGATIVE' : qty === 0 ? 'OUT OF STOCK' : qty <= reorder ? 'LOW STOCK' : 'OK'
    };
  });
}

function postTransaction(token, d) {
  d = d || {};
  const type = String(d.type || '');
  const perm = { RECEIVE: 'receive_stock', ISSUE: 'issue_stock', ADJUST: 'adjust_stock' }[type];
  if (!perm) throw new Error('Unsupported transaction type.');
  const s = require_(token, perm, false), ss = db_(), bid = s.businessId, loc = String(d.locationId || '');
  if (!Array.isArray(d.lines) || !d.lines.length) throw new Error('Add at least one item line.');
  if (d.lines.length > 100) throw new Error('Maximum of 100 lines per transaction.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!rows_(ss, 'Locations', bid).some(r => r[0] === loc && r[3] === 'Active')) throw new Error('Invalid location.');
    const items = rows_(ss, 'Items', bid), units = rows_(ss, 'Units', bid);
    const allowNeg = isTrue_(getSetting_(ss, bid, 'allow_negative_stock', 'TRUE'));
    const bal = balanceMap_(ss, bid), pending = {};

    // 1) I-validate LAHAT ng lines bago magsulat ng kahit ano (walang partial posting).
    const prepared = d.lines.map((line, n) => {
      const item = items.find(x => x[0] === line.itemId && x[8] === 'Active');
      if (!item) throw new Error('Line ' + (n + 1) + ': item not found in this business.');
      const qty = Number(line.qty);
      if (!(qty > 0)) throw new Error('Line ' + (n + 1) + ': quantity must be greater than zero.');
      const unit = String(line.unitId || item[5]);
      if (!units.some(u => u[0] === unit)) throw new Error('Line ' + (n + 1) + ': invalid unit.');
      const conv = conversion_(ss, bid, item[0], unit);
      const base = r6_(qty * conv);
      const cost = Math.max(0, num_(line.unitCost));
      let sign = type === 'RECEIVE' ? 1 : type === 'ISSUE' ? -1 : Number(line.adjustSign || 1);
      if (sign !== 1 && sign !== -1) throw new Error('Line ' + (n + 1) + ': invalid adjustment direction.');
      const expiry = String(line.expiryDate || '');
      if (expiry && !/^\d{4}-\d{2}-\d{2}$/.test(expiry)) throw new Error('Line ' + (n + 1) + ': invalid expiry date.');
      if (sign > 0 && isTrue_(item[7]) && !expiry) throw new Error('Line ' + (n + 1) + ': expiry date is required for ' + item[3] + '.');
      const delta = base * sign, key = loc + '|' + item[0];
      pending[key] = (pending[key] || 0) + delta;
      if (!allowNeg && (bal[key] || 0) + pending[key] < -1e-9) throw new Error('Insufficient stock for ' + item[3] + '. Negative on-hand is disabled.');
      return { item: item, qty: qty, unit: unit, conv: conv, base: base, delta: delta, cost: cost, baseCost: cost / conv,
        batch: String(line.batchNo || ''), expiry: expiry, notes: String(line.notes || '') };
    });

    const tid = id_('TXN'), stamp = now_();
    const canApprove = s.role === 'BusinessAdmin' || hasPerm_(s, 'approve_transactions');
    const status = (type === 'ADJUST' && !canApprove) ? 'PendingApproval' : 'Posted';
    ss.getSheetByName('Transactions').appendRow([tid, bid, type, String(d.reference || ''), loc, '', '', status, s.userId, stamp, String(d.notes || ''), '']);
    const lineSheet = ss.getSheetByName('TransactionLines');

    // 2) Isulat. Para sa ADJUST, signed ang QtyBase (para tama ang approval mamaya).
    prepared.forEach(p => {
      let costPerEntered = p.cost;
      if (status === 'Posted') {
        const res = applyMovement_(ss, bid, loc, p.item[0], p.delta, p.baseCost, tid, stamp, p.batch, p.expiry);
        if (p.delta < 0) costPerEntered = r6_(res.totalCost / p.qty);
      }
      lineSheet.appendRow([id_('LIN'), bid, tid, p.item[0], p.qty, p.unit, type === 'ADJUST' ? p.delta : p.base, costPerEntered, p.batch, p.expiry, p.notes]);
    });
    if (status === 'PendingApproval')
      ss.getSheetByName('Approvals').appendRow([id_('APR'), bid, tid, s.userId, '', 'Pending', 'Adjustment requires approval', stamp, '']);
    audit_(bid, s.userId, 'POST_TRANSACTION', 'Transaction', tid, JSON.stringify({ type: type, status: status }));
    return { ok: true, transactionId: tid, status: status,
      message: status === 'Posted' ? 'Transaction posted.' : 'Submitted for approval; stock unchanged until approval.' };
  } finally { lock.releaseLock(); }
}

function listPendingApprovals(token) {
  const s = require_(token, 'approve_transactions', false), ss = db_(), bid = s.businessId;
  const users = rows_(ss, 'Users', bid), items = rows_(ss, 'Items', bid), locs = rows_(ss, 'Locations', bid), lines = rows_(ss, 'TransactionLines', bid);
  return rows_(ss, 'Transactions', bid).filter(t => t[7] === 'PendingApproval').map(t => {
    const u = users.find(x => x[0] === t[8]), l = locs.find(x => x[0] === t[4]);
    const text = lines.filter(x => x[2] === t[0]).map(x => {
      const it = items.find(i => i[0] === x[3]);
      return (it ? it[3] : x[3]) + ' ' + (num_(x[6]) >= 0 ? '+' : '') + num_(x[6]);
    }).join('; ');
    return { id: t[0], type: t[2], reference: t[3], location: l ? l[2] : '', requestedBy: u ? u[3] : t[8], createdAt: t[9], notes: t[10], summary: text };
  });
}

function approveTransaction(token, transactionId, decision, reason) {
  const s = require_(token, 'approve_transactions', false), ss = db_(), bid = s.businessId;
  if (!['Approve', 'Reject'].includes(decision)) throw new Error('Invalid decision.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const txsh = ss.getSheetByName('Transactions'), td = txsh.getDataRange().getValues();
    const row = td.findIndex((r, i) => i > 0 && r[0] === transactionId && r[1] === bid);
    if (row < 1) throw new Error('Transaction not found.');
    const tx = td[row];
    if (tx[7] !== 'PendingApproval') throw new Error('Transaction is not pending approval.');
    if (tx[8] === s.userId && s.role !== 'BusinessAdmin') throw new Error('You cannot approve your own transaction.');
    const lines = ss.getSheetByName('TransactionLines').getDataRange().getValues().slice(1).filter(r => r[1] === bid && r[2] === transactionId);
    const stamp = now_();
    if (decision === 'Approve') {
      const allowNeg = isTrue_(getSetting_(ss, bid, 'allow_negative_stock', 'TRUE')), bal = balanceMap_(ss, bid), pending = {};
      lines.forEach(l => {
        const key = tx[4] + '|' + l[3];
        pending[key] = (pending[key] || 0) + num_(l[6]);
        if (!allowNeg && (bal[key] || 0) + pending[key] < -1e-9) throw new Error('Approval would cause negative stock, which is disabled.');
      });
      lines.forEach(l => {
        const delta = num_(l[6]), absBase = Math.abs(delta), qty = Math.abs(num_(l[4]));
        const baseCost = absBase > 0 ? num_(l[7]) * qty / absBase : 0;
        applyMovement_(ss, bid, tx[4], l[3], delta, baseCost, transactionId, stamp, String(l[8] || ''), String(l[9] || ''));
      });
    }
    txsh.getRange(row + 1, 8).setValue(decision === 'Approve' ? 'Posted' : 'Rejected');
    const sh = ss.getSheetByName('Approvals'), a = sh.getDataRange().getValues();
    for (let i = 1; i < a.length; i++) if (a[i][2] === transactionId && a[i][1] === bid && a[i][5] === 'Pending') {
      sh.getRange(i + 1, 5, 1, 5).setValues([[s.userId, decision === 'Approve' ? 'Approved' : 'Rejected', String(reason || ''), a[i][7], stamp]]);
      break;
    }
    audit_(bid, s.userId, decision.toUpperCase() + '_TRANSACTION', 'Transaction', transactionId, String(reason || ''));
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ============================ TRANSFERS ============================ */

function createTransfer(token, d) {
  const s = require_(token, 'transfer_stock', false), ss = db_(), bid = s.businessId;
  d = d || {};
  const locs = rows_(ss, 'Locations', bid).filter(r => r[3] === 'Active');
  if (!locs.some(r => r[0] === d.from) || !locs.some(r => r[0] === d.to) || d.from === d.to) throw new Error('Choose two different valid locations.');
  if (!Array.isArray(d.lines) || !d.lines.length) throw new Error('Add transfer lines.');
  const items = rows_(ss, 'Items', bid), units = rows_(ss, 'Units', bid);
  const prepared = d.lines.map((l, n) => {
    const item = items.find(x => x[0] === l.itemId && x[8] === 'Active');
    if (!item) throw new Error('Line ' + (n + 1) + ': item not found.');
    const qty = Number(l.qty);
    if (!(qty > 0)) throw new Error('Line ' + (n + 1) + ': quantity must be greater than zero.');
    const unit = String(l.unitId || item[5]);
    if (!units.some(u => u[0] === unit)) throw new Error('Line ' + (n + 1) + ': invalid unit.');
    return { item: item, qty: qty, unit: unit, base: r6_(qty * conversion_(ss, bid, item[0], unit)), batch: String(l.batchNo || ''), expiry: String(l.expiryDate || '') };
  });
  const tid = id_('TRF'), txid = id_('TXN'), stamp = now_();
  ss.getSheetByName('Transactions').appendRow([txid, bid, 'TRANSFER', String(d.reference || ''), '', d.from, d.to, 'Requested', s.userId, stamp, String(d.notes || ''), '']);
  ss.getSheetByName('Transfers').appendRow([tid, bid, txid, d.from, d.to, 'Requested', '', '', '']);
  const lsh = ss.getSheetByName('TransactionLines');
  prepared.forEach(p => lsh.appendRow([id_('LIN'), bid, txid, p.item[0], p.qty, p.unit, p.base, 0, p.batch, p.expiry, '']));
  audit_(bid, s.userId, 'CREATE_TRANSFER', 'Transfer', tid, 'Requested');
  return { ok: true, transferId: tid };
}

function listTransfers(token) {
  const s = requireAny_(token, ['transfer_stock', 'confirm_transfers', 'view_reports'], 'transfers'), ss = db_(), bid = s.businessId;
  const locs = rows_(ss, 'Locations', bid), items = rows_(ss, 'Items', bid), units = rows_(ss, 'Units', bid), lines = rows_(ss, 'TransactionLines', bid);
  const name = id => { const l = locs.find(x => x[0] === id); return l ? l[2] : id; };
  return rows_(ss, 'Transfers', bid).map(t => ({
    id: t[0], from: name(t[3]), to: name(t[4]), status: t[5], dispatchedAt: t[6], receivedAt: t[7], notes: t[8],
    summary: lines.filter(l => l[2] === t[2]).map(l => {
      const it = items.find(i => i[0] === l[3]), u = units.find(x => x[0] === l[5]);
      return (it ? it[3] : l[3]) + ' × ' + num_(l[4]) + ' ' + (u ? u[3] : '');
    }).join('; ')
  })).reverse();
}

function dispatchTransfer(token, transferId) {
  const s = require_(token, 'transfer_stock', false), ss = db_(), bid = s.businessId;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName('Transfers'), d = sh.getDataRange().getValues();
    const i = d.findIndex((r, n) => n > 0 && r[0] === transferId && r[1] === bid);
    if (i < 1 || d[i][5] !== 'Requested') throw new Error('Transfer not found or already dispatched.');
    const txid = d[i][2], from = d[i][3];
    const lsh = ss.getSheetByName('TransactionLines'), ld = lsh.getDataRange().getValues();
    const idx = []; ld.forEach((r, n) => { if (n > 0 && r[1] === bid && r[2] === txid) idx.push(n); });
    const allowNeg = isTrue_(getSetting_(ss, bid, 'allow_negative_stock', 'TRUE')), bal = balanceMap_(ss, bid), pending = {};
    idx.forEach(n => {
      const key = from + '|' + ld[n][3];
      pending[key] = (pending[key] || 0) - num_(ld[n][6]);
      if (!allowNeg && (bal[key] || 0) + pending[key] < -1e-9) throw new Error('Insufficient stock at source location.');
    });
    const stamp = now_();
    idx.forEach(n => {
      const r = ld[n], base = num_(r[6]);
      const res = applyMovement_(ss, bid, from, r[3], -base, 0, txid, stamp, '', '');
      // Itala ang FIFO cost na umalis para madala ito sa destination (UnitCost = cost per entered unit).
      lsh.getRange(n + 1, 8).setValue(num_(r[4]) > 0 ? r6_(res.totalCost / num_(r[4])) : 0);
    });
    sh.getRange(i + 1, 6, 1, 2).setValues([['InTransit', stamp]]);
    setTransactionStatus_(ss, bid, txid, 'InTransit');
    audit_(bid, s.userId, 'DISPATCH_TRANSFER', 'Transfer', transferId, 'InTransit');
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function receiveTransfer(token, transferId, notes) {
  const s = require_(token, 'confirm_transfers', false), ss = db_(), bid = s.businessId;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName('Transfers'), d = sh.getDataRange().getValues();
    const i = d.findIndex((r, n) => n > 0 && r[0] === transferId && r[1] === bid);
    if (i < 1 || d[i][5] !== 'InTransit') throw new Error('Transfer is not in transit.');
    const txid = d[i][2], to = d[i][4], stamp = now_();
    ss.getSheetByName('TransactionLines').getDataRange().getValues().slice(1).filter(r => r[1] === bid && r[2] === txid).forEach(l => {
      const base = num_(l[6]), qty = num_(l[4]);
      const baseCost = base > 0 ? num_(l[7]) * qty / base : 0;
      applyMovement_(ss, bid, to, l[3], base, baseCost, txid, stamp, String(l[8] || ''), String(l[9] || ''));
    });
    sh.getRange(i + 1, 6, 1, 4).setValues([['Received', d[i][6], stamp, String(notes || '')]]);
    setTransactionStatus_(ss, bid, txid, 'Received');
    audit_(bid, s.userId, 'RECEIVE_TRANSFER', 'Transfer', transferId, String(notes || ''));
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ============================ STOCK COUNTS ============================ */

function createStockCount(token, d) {
  const s = require_(token, 'stock_count', false), ss = db_(), bid = s.businessId;
  d = d || {};
  if (!rows_(ss, 'Locations', bid).some(r => r[0] === d.locationId)) throw new Error('Invalid location.');
  const items = rows_(ss, 'Items', bid).filter(i => i[8] === 'Active');
  if (!items.length) throw new Error('Add items first.');
  const bal = balanceMap_(ss, bid), id = id_('CNT');
  ss.getSheetByName('StockCounts').appendRow([id, bid, d.locationId, 'Open', s.userId, now_(), '', String(d.notes || '')]);
  const rows = items.map(it => [id_('CLN'), bid, id, it[0], r6_(bal[d.locationId + '|' + it[0]] || 0), '', '']);
  const sh = ss.getSheetByName('StockCountLines');
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  audit_(bid, s.userId, 'CREATE_COUNT', 'StockCount', id, d.locationId);
  return { ok: true, countId: id };
}

function listCounts(token) {
  const s = require_(token, 'stock_count', false), ss = db_(), locs = rows_(ss, 'Locations', s.businessId);
  return rows_(ss, 'StockCounts', s.businessId).map(c => {
    const l = locs.find(x => x[0] === c[2]);
    return { id: c[0], location: l ? l[2] : c[2], status: c[3], createdAt: c[5], postedAt: c[6], notes: c[7] };
  }).reverse();
}

function getCount(token, countId) {
  const s = require_(token, 'stock_count', false), ss = db_(), bid = s.businessId;
  const c = rows_(ss, 'StockCounts', bid).find(x => x[0] === countId);
  if (!c) throw new Error('Count not found.');
  const items = rows_(ss, 'Items', bid), units = rows_(ss, 'Units', bid), loc = rows_(ss, 'Locations', bid).find(x => x[0] === c[2]);
  const lines = rows_(ss, 'StockCountLines', bid).filter(l => l[2] === countId).map(l => {
    const it = items.find(i => i[0] === l[3]), u = it && units.find(x => x[0] === it[5]);
    return { itemId: l[3], sku: it ? it[2] : '', name: it ? it[3] : l[3], unit: u ? u[3] : '', expected: num_(l[4]),
      counted: l[5] === '' || l[5] == null ? null : num_(l[5]) };
  });
  return { id: c[0], status: c[3], location: loc ? loc[2] : c[2], notes: c[7], lines: lines };
}

function saveCountLines(token, countId, lines) {
  const s = require_(token, 'stock_count', false), ss = db_(), bid = s.businessId;
  if (!Array.isArray(lines)) throw new Error('Invalid data.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!rows_(ss, 'StockCounts', bid).some(c => c[0] === countId && c[3] === 'Open')) throw new Error('Count is not open.');
    const sh = ss.getSheetByName('StockCountLines'), d = sh.getDataRange().getValues();
    let saved = 0;
    lines.forEach(x => {
      if (x.countedQty === '' || x.countedQty == null) return;
      const n = Number(x.countedQty);
      if (!Number.isFinite(n) || n < 0) throw new Error('Invalid count quantity.');
      for (let i = 1; i < d.length; i++) if (d[i][1] === bid && d[i][2] === countId && d[i][3] === x.itemId) {
        sh.getRange(i + 1, 6, 1, 2).setValues([[n, r6_(n - num_(d[i][4]))]]);
        saved++;
        break;
      }
    });
    return { ok: true, saved: saved };
  } finally { lock.releaseLock(); }
}

function postStockCount(token, countId) {
  const s = require_(token, 'stock_count', false), ss = db_(), bid = s.businessId;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = ss.getSheetByName('StockCounts'), d = sh.getDataRange().getValues();
    const i = d.findIndex((r, n) => n > 0 && r[0] === countId && r[1] === bid);
    if (i < 1 || d[i][3] !== 'Open') throw new Error('Count is not open.');
    const loc = d[i][2];
    const lines = rows_(ss, 'StockCountLines', bid).filter(r => r[2] === countId);
    if (lines.some(l => l[5] === '' || l[5] == null)) throw new Error('Count every item before posting.');
    const layers = rows_(ss, 'FIFOLayers', bid), stamp = now_();
    lines.forEach(l => {
      const delta = r6_(num_(l[5]) - num_(l[4]));
      if (!delta) return;
      // Para sa sobra, gamitin ang pinakahuling known cost ng item sa location na ito.
      const last = layers.filter(x => x[2] === loc && x[3] === l[3]).pop();
      applyMovement_(ss, bid, loc, l[3], delta, last ? num_(last[8]) : 0, countId, stamp, '', '');
    });
    sh.getRange(i + 1, 4).setValue('Posted');
    sh.getRange(i + 1, 7).setValue(stamp);
    audit_(bid, s.userId, 'POST_STOCK_COUNT', 'StockCount', countId, 'Posted');
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function cancelStockCount(token, countId) {
  const s = require_(token, 'stock_count', false), ss = db_(), bid = s.businessId;
  const sh = ss.getSheetByName('StockCounts'), d = sh.getDataRange().getValues();
  const i = d.findIndex((r, n) => n > 0 && r[0] === countId && r[1] === bid);
  if (i < 1 || d[i][3] !== 'Open') throw new Error('Count is not open.');
  sh.getRange(i + 1, 4).setValue('Cancelled');
  audit_(bid, s.userId, 'CANCEL_STOCK_COUNT', 'StockCount', countId, '');
  return { ok: true };
}

/* ============================ REPORTS ============================ */

function reportTransactions(token, fromDate, toDate) {
  const s = require_(token, 'view_reports', false), ss = db_(), bid = s.businessId;
  const users = rows_(ss, 'Users', bid), locs = rows_(ss, 'Locations', bid);
  const ln = id => { const l = locs.find(x => x[0] === id); return l ? l[2] : ''; };
  return rows_(ss, 'Transactions', bid).filter(r => {
    const day = String(r[9]).slice(0, 10);
    return (!fromDate || day >= fromDate) && (!toDate || day <= toDate);
  }).map(r => {
    const u = users.find(x => x[0] === r[8]);
    return { id: r[0], type: r[2], reference: r[3], location: r[4] ? ln(r[4]) : (ln(r[5]) + ' → ' + ln(r[6])), status: r[7],
      createdBy: u ? u[2] : r[8], createdAt: r[9], notes: r[10] };
  });
}

function exportInventoryCsv(token) {
  const list = inventory(token, '');
  const rows = [['ItemID', 'SKU', 'Name', 'Category', 'Unit', 'Qty', 'Status', 'Valuation']]
    .concat(list.map(x => [x.itemId, x.sku, x.name, x.category, x.unit, x.qty, x.status, x.valuation === null ? '' : x.valuation]));
  return rows.map(r => r.map(csv_).join(',')).join('\r\n');
}
function csv_(v) {
  let t = String(v == null ? '' : v);
  if (/^[=+\-@\t\r]/.test(t)) t = "'" + t;   // iwas CSV/formula injection
  return '"' + t.replace(/"/g, '""') + '"';
}

/* ============================ INTERNAL HELPERS ============================ */

/** Basahin ang rows ng business. Hinahanap ang BusinessID column ayon sa header (hindi laging column 2). */
function rows_(ss, sheet, bid) {
  const sh = ss.getSheetByName(sheet);
  if (!sh || sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues(), bi = d[0].indexOf('BusinessID');
  return d.slice(1).filter(r => r[bi] === bid).map(r => r.map(cell_));
}
/** Hindi puwedeng ibalik ang Date object sa google.script.run, kaya gawing string. */
function cell_(v) { return v instanceof Date ? Utilities.formatDate(v, U.TZ, "yyyy-MM-dd'T'HH:mm:ss") : v; }

function balanceMap_(ss, bid) {
  const m = {};
  rows_(ss, 'StockBalances', bid).forEach(r => { m[r[1] + '|' + r[2]] = num_(r[3]); });
  return m;
}

function getSetting_(ss, bid, key, def) {
  const d = ss.getSheetByName('Settings').getDataRange().getValues().slice(1);
  const own = d.find(r => r[0] === bid && r[1] === key), glob = d.find(r => r[0] === '*' && r[1] === key);
  return own ? String(own[2]) : glob ? String(glob[2]) : def;
}

function updateBalance_(ss, bid, loc, item, delta) {
  const sh = ss.getSheetByName('StockBalances'), d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) if (d[i][0] === bid && d[i][1] === loc && d[i][2] === item) {
    sh.getRange(i + 1, 4, 1, 2).setValues([[r6_(num_(d[i][3]) + delta), now_()]]);
    return;
  }
  sh.appendRow([bid, loc, item, r6_(delta), now_()]);
}

/** FIFO: bawasan ang pinakamatandang layers. Ibabalik ang total cost na nakonsumo. */
function consumeFIFO_(ss, bid, loc, item, qty) {
  const sh = ss.getSheetByName('FIFOLayers'), d = sh.getDataRange().getValues();
  let remain = qty, totalCost = 0;
  for (let i = 1; i < d.length && remain > 1e-9; i++) {
    if (d[i][1] === bid && d[i][2] === loc && d[i][3] === item && num_(d[i][7]) > 0) {
      const take = Math.min(remain, num_(d[i][7]));
      sh.getRange(i + 1, 8).setValue(r6_(num_(d[i][7]) - take));
      totalCost += take * num_(d[i][8]);
      remain -= take;
    }
  }
  // Kung may natirang remain: negative on-hand (walang cost layer para sa shortage).
  return { totalCost: totalCost, shortage: Math.max(0, remain) };
}

/** Isang pinto para sa lahat ng stock movement: balance + FIFO layer. */
function applyMovement_(ss, bid, loc, itemId, delta, baseCost, txId, stamp, batch, expiry) {
  updateBalance_(ss, bid, loc, itemId, delta);
  if (delta > 0) {
    ss.getSheetByName('FIFOLayers').appendRow([id_('FIFO'), bid, loc, itemId, txId, stamp, delta, delta, r6_(baseCost), batch || '', expiry || '']);
    return { totalCost: delta * baseCost, shortage: 0 };
  }
  if (delta < 0) return consumeFIFO_(ss, bid, loc, itemId, -delta);
  return { totalCost: 0, shortage: 0 };
}

function conversion_(ss, bid, itemId, unit) {
  const it = rows_(ss, 'Items', bid).find(r => r[0] === itemId);
  if (!it) throw new Error('Item not found in this business.');
  if (it[5] === unit) return 1;
  const r = rows_(ss, 'ItemUnits', bid).find(x => x[1] === itemId && x[2] === unit);
  if (!r) throw new Error('Unit conversion not configured for ' + it[3] + '. Add it in Items > Unit conversions.');
  const n = Number(r[3]);
  if (!(n > 0)) throw new Error('Invalid unit conversion.');
  return n;
}

function setTransactionStatus_(ss, bid, id, status) {
  const sh = ss.getSheetByName('Transactions'), d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) if (d[i][0] === id && d[i][1] === bid) { sh.getRange(i + 1, 8).setValue(status); return; }
}

function audit_(bid, uid, action, entity, eid, details) {
  try {
    db_().getSheetByName('AuditLog').appendRow([id_('AUD'), bid, uid, action, entity, eid, String(details || ''), now_()]);
  } catch (e) { console.error(e); }
}

/** SHA-256 (hex) with salt. Hindi ito KDF; para sa production, gumamit ng tunay na auth provider. */
function hash_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + password, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function validatePassword_(p) {
  p = String(p || '');
  if (p.length < 10) throw new Error('Password must be at least 10 characters.');
  if (!/[A-Za-z]/.test(p) || !/[0-9]/.test(p)) throw new Error('Password must contain letters and numbers.');
}
function clean_(s) { return String(s || '').trim().replace(/[^A-Za-z0-9._-]/g, ''); }
function id_(prefix) { return prefix + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase(); }
function now_() { return Utilities.formatDate(new Date(), U.TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function isTrue_(v) { return v === true || String(v).toLowerCase() === 'true'; }
function num_(v) { const n = Number(v); return Number.isFinite(n) ? n : 0; }
function r6_(n) { return Math.round(n * 1e6) / 1e6; }
