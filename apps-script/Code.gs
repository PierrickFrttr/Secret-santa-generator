/**
 * Secret Santa — serveur (Google Apps Script)
 * Base de données : une feuille Google créée automatiquement au premier appel.
 * Déploiement : Déployer > Nouveau déploiement > Application Web
 *   Exécuter en tant que : moi — Accès : Tout le monde
 */
const PROPS = PropertiesService.getScriptProperties();
const T_DRAWS = 'Tirages', T_CLAIMS = 'Découvertes';

function db_() {
  let id = PROPS.getProperty('SHEET_ID'), ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Secret Santa – base de données');
    PROPS.setProperty('SHEET_ID', ss.getId());
  }
  return ss;
}
function sheet_(name) {
  const ss = db_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    const head = name === T_DRAWS
      ? ['id', 'titre', 'budget', 'date', 'prenoms', 'cercle (codé)', 'cree_le']
      : ['id', 'index', 'prenom', 'decouvert_le'];
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:G').setNumberFormat('@');
    const def = ss.getSheetByName('Feuille 1') || ss.getSheetByName('Sheet1');
    if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  }
  return sh;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e) { return handle_((e && e.parameter) || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (x) {}
  return handle_(p);
}

function handle_(p) {
  try {
    switch (p.action) {
      case 'get': return out_(getDraw_(String(p.id || '')));
      case 'claim': return out_(claim_(String(p.id || ''), Number(p.idx)));
      case 'status': return out_({ ok: true, setup: !!PROPS.getProperty('ADMIN_HASH') });
      case 'setup': return out_(setup_(String(p.pass || '')));
    }
    if (!isAdmin_(String(p.pass || ''))) return out_({ error: 'auth' });
    switch (p.action) {
      case 'login': return out_({ ok: true });
      case 'list': return out_({ ok: true, draws: list_() });
      case 'create': return out_(create_(p));
      case 'reset': return out_(reset_(String(p.id || ''), Number(p.idx)));
      case 'delete': return out_(delete_(String(p.id || '')));
    }
    return out_({ error: 'action' });
  } catch (err) {
    return out_({ error: String((err && err.message) || err) });
  }
}

/* ---------- organisateur ---------- */
function hash_(pass) {
  const salt = PROPS.getProperty('SALT');
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pass, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(d);
}
function setup_(pass) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (PROPS.getProperty('ADMIN_HASH')) return { error: 'already' };
    if (pass.length < 8) return { error: 'short' };
    PROPS.setProperty('SALT', Utilities.getUuid());
    PROPS.setProperty('ADMIN_HASH', hash_(pass));
    db_();
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function isAdmin_(pass) {
  const h = PROPS.getProperty('ADMIN_HASH');
  if (!h || !pass) return false;
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= 10) throw new Error('locked');
  const ok = hash_(pass) === h;
  if (!ok) cache.put('fails', String(fails + 1), 600);
  return ok;
}

/* ---------- données ---------- */
function rows_(name) {
  const v = sheet_(name).getDataRange().getValues();
  v.shift();
  return v;
}
function findDraw_(id) {
  const v = rows_(T_DRAWS);
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === id) return { row: i + 2, r: v[i] };
  return null;
}
function drawObj_(r) {
  return { id: String(r[0]), title: String(r[1]), budget: String(r[2]), date: String(r[3]), names: JSON.parse(r[4]), created: String(r[6]) };
}
function claimsOf_(id) {
  return rows_(T_CLAIMS).filter(c => String(c[0]) === id).map(c => ({ idx: Number(c[1]), at: String(c[3]) }));
}
function cycleOf_(r) {
  return JSON.parse(Utilities.newBlob(Utilities.base64Decode(String(r[5]))).getDataAsString());
}

function getDraw_(id) {
  const f = findDraw_(id);
  if (!f) return { error: 'notfound' };
  const d = drawObj_(f.r);
  d.claimed = claimsOf_(id).map(c => c.idx);
  return { ok: true, draw: d };
}

function claim_(id, idx) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const f = findDraw_(id);
    if (!f) return { error: 'notfound' };
    const d = drawObj_(f.r);
    if (!(idx >= 0 && idx < d.names.length)) return { error: 'idx' };
    if (claimsOf_(id).some(c => c.idx === idx)) return { error: 'taken' };
    sheet_(T_CLAIMS).appendRow([id, String(idx), d.names[idx], new Date().toISOString()]);
    const c = cycleOf_(f.r);
    return { ok: true, target: d.names[c[idx]] };
  } finally { lock.releaseLock(); }
}

function slug_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 32).replace(/-+$/, '');
}
function create_(p) {
  const clean = s => String(s || '').trim().replace(/\s+/g, ' ');
  const names = (Array.isArray(p.names) ? p.names : []).map(n => clean(n).slice(0, 30)).filter(Boolean);
  const low = names.map(n => n.toLowerCase());
  if (names.length < 3) return { error: 'few' };
  if (names.length > 60) return { error: 'many' };
  if (new Set(low).size !== low.length) return { error: 'dup' };
  const title = clean(p.title).slice(0, 50), budget = clean(p.budget).slice(0, 20), date = clean(p.date).slice(0, 40);
  // Sattolo : un seul cercle, personne ne se tire soi-même
  const c = names.map((_, i) => i);
  for (let i = c.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * i); const t = c[i]; c[i] = c[j]; c[j] = t; }
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const taken = new Set(rows_(T_DRAWS).map(r => String(r[0])));
    let base = slug_(p.slug) || slug_(title) || 'tirage', id = base, k = 2;
    if (['admin', 'index', '404'].indexOf(base) >= 0) id = base = base + '-noel';
    while (taken.has(id)) id = base + '-' + (k++);
    const enc = Utilities.base64Encode(JSON.stringify(c), Utilities.Charset.UTF_8);
    sheet_(T_DRAWS).appendRow([id, title, budget, date, JSON.stringify(names), enc, new Date().toISOString()]);
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}
function list_() {
  const claims = rows_(T_CLAIMS);
  return rows_(T_DRAWS).map(r => {
    const d = drawObj_(r);
    d.claimed = claims.filter(c => String(c[0]) === d.id).map(c => ({ idx: Number(c[1]), at: String(c[3]) }));
    return d;
  }).reverse();
}
function reset_(id, idx) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = sheet_(T_CLAIMS), v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === id && Number(v[i][1]) === idx) sh.deleteRow(i + 1);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function delete_(id) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const f = findDraw_(id);
    if (f) sheet_(T_DRAWS).deleteRow(f.row);
    const sh = sheet_(T_CLAIMS), v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === id) sh.deleteRow(i + 1);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* À lancer une fois depuis l'éditeur pour autoriser l'accès à Google Sheets */
function initialiser() { sheet_(T_DRAWS); sheet_(T_CLAIMS); Logger.log(db_().getUrl()); }
