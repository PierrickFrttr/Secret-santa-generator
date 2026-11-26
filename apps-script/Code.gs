/**
 * Secret Santa — serveur (Google Apps Script)
 * Base de données : une feuille Google créée automatiquement au premier appel.
 * Déploiement : Déployer > Nouveau déploiement > Application Web
 *   Exécuter en tant que : moi — Accès : Tout le monde
 */
const PROPS = PropertiesService.getScriptProperties();
const T_DRAWS = 'Tirages', T_CLAIMS = 'Découvertes';
const HEAD_DRAWS = ['id', 'titre', 'budget', 'date', 'prenoms', 'cercle (codé)', 'cree_le', 'mode', 'codes', 'organisateur'];
const T_ORGS = 'Organisateurs', HEAD_ORGS = ['identifiant', 'sel', 'empreinte', 'cree_le'];
const SUPER = 'pierrick';

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
    const head = name === T_DRAWS ? HEAD_DRAWS : name === T_ORGS ? HEAD_ORGS : ['id', 'index', 'prenom', 'decouvert_le'];
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:J').setNumberFormat('@');
    const def = ss.getSheetByName('Feuille 1') || ss.getSheetByName('Sheet1');
    if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  }
  if (name === T_DRAWS && sh.getLastColumn() < HEAD_DRAWS.length) {
    sh.getRange(1, 1, 1, HEAD_DRAWS.length).setValues([HEAD_DRAWS]).setFontWeight('bold');
    sh.getRange('A:J').setNumberFormat('@');
    sh.getRange('A:J').setNumberFormat('@');
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
      case 'get': return out_(getDraw_(String(p.id || ''), String(p.code || '')));
      case 'claim': return out_(claim_(String(p.id || ''), Number(p.idx), String(p.code || '')));
      case 'status': return out_({ ok: true, setup: !!PROPS.getProperty('ADMIN_HASH') });
      case 'setup': return out_(setup_(String(p.pass || '')));
      case 'join': return out_(join_(p));
      case 'invitecheck': return out_({ ok: !!invites_()[String(p.code || '')] });
    }
    const me = auth_(p);
    if (!me) return out_({ error: 'auth' });
    const id = String(p.id || '');
    if (['reset', 'delete', 'verify'].indexOf(p.action) >= 0 && !owns_(id, me)) return out_({ error: 'notfound' });
    switch (p.action) {
      case 'login': return out_({ ok: true, user: me.id, super: me.super });
      case 'list': return out_({ ok: true, draws: list_(me) });
      case 'create': return out_(create_(p, me));
      case 'reset': return out_(reset_(id, Number(p.idx)));
      case 'delete': return out_(delete_(id));
      case 'verify': return out_(verify_(id));
    }
    if (!me.super) return out_({ error: 'forbidden' });
    switch (p.action) {
      case 'orgs': return out_(orgs_());
      case 'invite': return out_(invite_());
      case 'delinvite': return out_(delInvite_(String(p.code || '')));
      case 'delorg': return out_(delOrg_(slug_(p.org)));
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
function hashWith_(salt, pass) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pass, Utilities.Charset.UTF_8));
}
function findOrg_(u) {
  const v = rows_(T_ORGS);
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === u) return { row: i + 2, r: v[i] };
  return null;
}
function auth_(p) {
  const u = slug_(p.user) || SUPER, pass = String(p.pass || '');
  if (!pass) return null;
  const cache = CacheService.getScriptCache(), key = 'fails:' + u;
  const fails = Number(cache.get(key) || 0);
  if (fails >= 10) throw new Error('locked');
  let ok = false;
  if (u === SUPER) { const h = PROPS.getProperty('ADMIN_HASH'); ok = !!h && hash_(pass) === h; }
  else { const o = findOrg_(u); ok = !!o && hashWith_(String(o.r[1]), pass) === String(o.r[2]); }
  if (!ok) { cache.put(key, String(fails + 1), 600); return null; }
  return { id: u, super: u === SUPER };
}
function ownerOf_(r) { return String(r[9] || '') || SUPER; }
function owns_(id, me) { const f = findDraw_(id); return !!f && ownerOf_(f.r) === me.id; }

/* ---------- invitations et organisateurs (réservé à pierrick) ---------- */
function invites_() {
  const all = JSON.parse(PROPS.getProperty('INVITES') || '{}'), now = Date.now(), out = {};
  Object.keys(all).forEach(k => { if (now - new Date(all[k]).getTime() < 14 * 864e5) out[k] = all[k]; });
  return out;
}
function invite_() {
  const AL = 'abcdefghjkmnpqrstuvwxyz23456789'; let code = '';
  for (let i = 0; i < 10; i++) code += AL[Math.floor(Math.random() * AL.length)];
  const all = invites_(); all[code] = new Date().toISOString();
  PROPS.setProperty('INVITES', JSON.stringify(all));
  return { ok: true, code: code };
}
function delInvite_(code) {
  const all = invites_(); delete all[code];
  PROPS.setProperty('INVITES', JSON.stringify(all));
  return { ok: true };
}
function join_(p) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const code = String(p.code || ''), all = invites_();
    if (!all[code]) return { error: 'invite' };
    const u = slug_(p.user), pass = String(p.pass || '');
    if (u.length < 2) return { error: 'user' };
    if (u === SUPER || u === 'admin' || findOrg_(u)) return { error: 'taken' };
    if (pass.length < 8) return { error: 'short' };
    const salt = Utilities.getUuid();
    sheet_(T_ORGS).appendRow([u, salt, hashWith_(salt, pass), new Date().toISOString()]);
    delete all[code]; PROPS.setProperty('INVITES', JSON.stringify(all));
    return { ok: true, user: u };
  } finally { lock.releaseLock(); }
}
function orgs_() {
  const draws = rows_(T_DRAWS), count = u => draws.filter(r => ownerOf_(r) === u).length;
  const list = rows_(T_ORGS).map(r => ({ id: String(r[0]), created: String(r[3]), draws: count(String(r[0])) }));
  const inv = invites_();
  return { ok: true, orgs: list, me: { id: SUPER, draws: count(SUPER) }, invites: Object.keys(inv).map(k => ({ code: k, created: inv[k] })) };
}
function delOrg_(u) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const o = findOrg_(u);
    if (!o) return { error: 'notfound' };
    sheet_(T_ORGS).deleteRow(o.row);
    const sh = sheet_(T_DRAWS), v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (String(v[i][9] || '') === u) sh.getRange(i + 1, 10).setValue(SUPER);
    return { ok: true };
  } finally { lock.releaseLock(); }
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
  return { id: String(r[0]), title: String(r[1]), budget: String(r[2]), date: String(r[3]), names: JSON.parse(r[4]), created: String(r[6]),
    mode: String(r[7] || '') === 'perso' ? 'perso' : 'commun', codes: r[8] ? JSON.parse(r[8]) : [] };
}
function claimsOf_(id) {
  return rows_(T_CLAIMS).filter(c => String(c[0]) === id).map(c => ({ idx: Number(c[1]), at: String(c[3]) }));
}
function cycleOf_(r) {
  return JSON.parse(Utilities.newBlob(Utilities.base64Decode(String(r[5]))).getDataAsString());
}

function codeIdx_(d, code) {
  const c = String(code || '').toLowerCase().split('-').pop();
  return c ? d.codes.indexOf(c) : -1;
}
function getDraw_(id, code) {
  const f = findDraw_(id);
  if (!f) return { error: 'notfound' };
  const d = drawObj_(f.r), claimed = claimsOf_(id).map(c => c.idx);
  const pub = { id: d.id, title: d.title, budget: d.budget, date: d.date, mode: d.mode };
  if (d.mode === 'perso') {
    const i = codeIdx_(d, code);
    if (i < 0) return { error: 'personal' };
    pub.idx = i; pub.me = d.names[i]; pub.opened = claimed.indexOf(i) >= 0; pub.count = d.names.length;
    return { ok: true, draw: pub };
  }
  pub.names = d.names; pub.claimed = claimed;
  return { ok: true, draw: pub };
}

function claim_(id, idx, code) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const f = findDraw_(id);
    if (!f) return { error: 'notfound' };
    const d = drawObj_(f.r);
    if (d.mode === 'perso') {
      idx = codeIdx_(d, code);
      if (idx < 0) return { error: 'personal' };
      // lien personnel : la personne peut revoir son résultat, on ne note que la première ouverture
      if (!claimsOf_(id).some(c => c.idx === idx)) sheet_(T_CLAIMS).appendRow([id, String(idx), d.names[idx], new Date().toISOString()]);
      return { ok: true, target: d.names[cycleOf_(f.r)[idx]] };
    }
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
function create_(p, me) {
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
    const perso = p.mode === 'perso', AL = 'abcdefghjkmnpqrstuvwxyz23456789', used = {};
    const codes = perso ? names.map(() => { let k; do { k = ''; for (let i = 0; i < 5; i++) k += AL[Math.floor(Math.random() * AL.length)]; } while (used[k]); used[k] = 1; return k; }) : [];
    sheet_(T_DRAWS).appendRow([id, title, budget, date, JSON.stringify(names), enc, new Date().toISOString(), perso ? 'perso' : 'commun', JSON.stringify(codes), me.id]);
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}
function list_(me) {
  const claims = rows_(T_CLAIMS);
  return rows_(T_DRAWS).filter(r => ownerOf_(r) === me.id).map(r => {
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
    // Nouveau résultat : on échange la cible de idx avec celle d'une personne qui n'a pas encore découvert la sienne
    const f = findDraw_(id);
    if (!f) return { ok: true, changed: false };
    const c = cycleOf_(f.r), n = c.length, t = c[idx];
    const opened = new Set(claimsOf_(id).map(x => x.idx));
    const cand = [];
    for (let j = 0; j < n; j++) if (j !== idx && !opened.has(j) && c[j] !== idx && t !== j) cand.push(j);
    if (!cand.length) return { ok: true, changed: false };
    const j = cand[Math.floor(Math.random() * cand.length)];
    c[idx] = c[j]; c[j] = t;
    sheet_(T_DRAWS).getRange(f.row, 6).setValue(Utilities.base64Encode(JSON.stringify(c), Utilities.Charset.UTF_8));
    return { ok: true, changed: true };
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

function verify_(id) {
  const f = findDraw_(id);
  if (!f) return { error: 'notfound' };
  const n = drawObj_(f.r).names.length, c = cycleOf_(f.r);
  const okLen = c.length === n;
  const recv = new Set(c);
  const once = okLen && recv.size === n && c.every(v => v >= 0 && v < n);
  const noSelf = c.every((v, i) => v !== i);
  let circles = 0;
  if (once) { const seen = new Array(n).fill(false); for (let s = 0; s < n; s++) if (!seen[s]) { circles++; let k = s; while (!seen[k]) { seen[k] = true; k = c[k]; } } }
  return { ok: true, n: n, circles: circles, checks: { everyoneGives: okLen, everyoneReceivesOnce: once, nobodySelf: noSelf, singleCircle: circles === 1 } };
}

/* À lancer une fois depuis l'éditeur pour autoriser l'accès à Google Sheets */
function initialiser() { sheet_(T_DRAWS); sheet_(T_CLAIMS); Logger.log(db_().getUrl()); }
