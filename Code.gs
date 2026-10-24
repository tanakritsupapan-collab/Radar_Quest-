/**
 * Radar Quest Game — Google Apps Script back end
 * Sheets: Users, Sessions, Ref (SAP), Updates (finds), NewItems, Locations, Points, History
 * - Ref is reference data from SAP. Players never write to it.
 * - Every find goes to Updates; things not in SAP go to NewItems (admin approves).
 *
 * First time: upload Asset_Master_SAP.xlsx and radar-quest-how-to-earn.gif to Google Drive,
 * run setup(), then importSAP() from the editor. Deploy as a web app.
 */

const TZ = 'Asia/Bangkok';
const TS_FMT = 'yyyy-MM-dd HH:mm:ss';
const DAY = 86400000;
const SESSION_DAYS = 30;
const PTS_MOVE = 10;          // moved to the ASSET ZONE
const PTS_BIG = 5;            // too big to move
const USER_RE = /^[a-z0-9._]{3,20}$/;
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const STATUSES = ['Used', 'Spare', 'Dispose'];
const START_LOCATIONS = ['ตึกปฏิบัติการ ชั้น 1 ห้อง 112', 'ตึกปฏิบัติการ ชั้น B ห้อง store เก่า'];
// Page players open (Google Sites page that embeds this web app). Empty = the web app link.
const ENTRY_URL = '';

const HEADERS = {
  Users:     ['Username','Email','Avatar','Avatar Img','Role','Status','Salt','PwHash','Act Token','Reset Token','Reset Expires','Created','Activated At','Reviewed By','Reviewed At','Points'],
  Sessions:  ['Token','Username','Expires'],
  Ref:       ['Asset13','Invent no','Asset Sup','Detail','Serial','Status','Owner','Admin Note','Removed','Added By','Updated By','Updated At'],
  Updates:   ['ID','Timestamp','Username','Asset13','Serial','Invent no','Detail','Location','Too Big','Barcode Check','Barcode Read','Points','State','Photo Barcode','Photo Device','Photo Wide','Thumb1','Thumb2','Thumb3',
             'Pick','Photo Age','Photo Hash','Flags','Check','Checked By','Checked At'],
  NewItems:  ['ID','Timestamp','Username','Number','Detail','Note','Location','Too Big','Barcode Read','State','Division','Remark','Status','Owner','Points','Reviewed By','Reviewed At','Photo Barcode','Photo Device','Photo Wide','Thumb1','Thumb2','Thumb3'],
  Locations: ['Location','New','Added By','Added At','Active'],
  Points:    ['Timestamp','Username','Delta','Reason','Ref','By'],
  History:   ['Timestamp','By','Role','Table','Key','Field','Old','New','Note'],
};

/* ===================== WEB APP ===================== */
function doGet(e){
  const p = (e && e.parameter) || {};
  let boot = {};
  try {
    if (p.act) boot.act = activate_(p.act);
    if (p.reset) boot.reset = resetTokenOk_(p.reset) ? p.reset : 'BAD';
  } catch (err){ boot.err = String(err.message || err); }
  const t = HtmlService.createTemplateFromFile('Index');
  t.boot = JSON.stringify(boot);
  return t.evaluate().setTitle('Radar Quest')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);   // allow the Google Sites entry page to embed it
}

/** google.script.run.api(action, token, payload) */
function api(action, token, payload){
  try {
    payload = payload || {};
    schema_();
    switch (action){
      case 'register':   return withLock_(() => register_(payload));
      case 'checkName':  return { free:!findUser_(String(payload.username || '').toLowerCase()) };
      case 'login':      return login_(payload);
      case 'forgot':     return forgot_(payload);
      case 'setNewPw':   return withLock_(() => setNewPw_(payload));
      case 'resend':     return withLock_(() => resend_(payload));
    }
    const me = auth_(token);
    if (!me) return { error:'AUTH' };
    const admin = me.Role === 'admin';
    const needAdmin = () => { if (!admin) throw new Error('Admin only'); };
    switch (action){
      case 'bootstrap':   return bootstrap_(me);
      case 'logout':      return logout_(token);
      case 'search':      return search_(payload);
      case 'lookup':      return lookup_(payload);
      case 'identify':    return identify_(payload);
      case 'submitFind':  return once_(payload.rid, () => submitFind_(me, payload));
      case 'submitNew':   return once_(payload.rid, () => submitNew_(me, payload));
      case 'photo':       return photo_(me, payload);
      case 'editNew':     return withLock_(() => editNew_(me, payload));
      case 'addLocation': return withLock_(() => addLocation_(me, payload));
      case 'myFinds':     return myFinds_(me);
      case 'setAvatar':   return withLock_(() => setAvatar_(me, payload));
      case 'changePw':    return withLock_(() => changePw_(me, payload));
      case 'guideGif':    return { gif:guideGif_() };
      // admin
      case 'adminHome':   needAdmin(); return adminHome_();
      case 'members':     needAdmin(); return members_();
      case 'setUser':     needAdmin(); return withLock_(() => setUser_(me, payload));
      case 'reviewNew':   needAdmin(); return withLock_(() => reviewNew_(me, payload));
      case 'finds':       needAdmin(); return finds_(payload);
      case 'checks':      needAdmin(); return checks_();
      case 'setCheck':    needAdmin(); return withLock_(() => setCheck_(me, payload));
      case 'asset':       needAdmin(); return assetAdmin_(payload);
      case 'editAsset':   needAdmin(); return withLock_(() => editAsset_(me, payload));
      case 'addAsset':    needAdmin(); return withLock_(() => addAsset_(me, payload));
      case 'adjustPoints':needAdmin(); return withLock_(() => adjustPoints_(me, payload));
      case 'resetFind':   needAdmin(); return withLock_(() => resetFind_(me, payload));
      case 'editLocation':needAdmin(); return withLock_(() => editLocation_(me, payload));
      case 'exportZip':   needAdmin(); return exportZip_(payload);
    }
    return { error:'Unknown action' };
  } catch (e){
    console.error(action, e && e.stack || e);
    return { error:String(e && e.message || e) };
  }
}
function withLock_(fn){ const l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }
/** Add columns that newer versions need to existing sheets (once). */
const SCHEMA = '3';
function schema_(force){
  const props = PropertiesService.getScriptProperties();
  if (!force && props.getProperty('SCHEMA') === SCHEMA) return;
  const ss = ss_();
  Object.keys(HEADERS).forEach(name => {
    const sh = ss.getSheetByName(name); if (!sh || sh.getLastRow() === 0) return;
    const have = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
    const miss = HEADERS[name].filter(h => !have.includes(h));
    if (!miss.length) return;
    if (sh.getMaxColumns() < have.length + miss.length) sh.insertColumnsAfter(sh.getMaxColumns(), have.length + miss.length - sh.getMaxColumns());
    sh.getRange(1, have.length + 1, sh.getMaxRows(), miss.length).setNumberFormat('@');
    sh.getRange(1, have.length + 1, 1, miss.length).setValues([miss]).setFontWeight('bold').setBackground('#24306E').setFontColor('#FFFFFF');
  });
  props.setProperty('SCHEMA', SCHEMA);
}
/** Only some columns of a sheet (Updates has big thumbnail columns). */
function cols_(name, want){
  const sh = ss_().getSheetByName(name), last = sh.getLastRow();
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  if (last < 2) return [];
  const idx = want.map(k => h.indexOf(k)), maxc = Math.max.apply(null, idx) + 1;
  return sh.getRange(2, 1, last - 1, maxc).getValues().map((r, i) => { const o = { _row:i + 2 }; want.forEach((k, j) => o[k] = idx[j] >= 0 ? cell_(r[idx[j]]) : ''); return o; });
}

/** Run a submit once per client request id, so a retry after a lost reply does not save twice. */
function once_(rid, fn){
  const cache = CacheService.getScriptCache(), key = rid ? 'rid:' + String(rid).slice(0, 60) : '';
  if (key){ const hit = cache.get(key); if (hit) return JSON.parse(hit); }
  const out = fn();
  if (key && out && !out.error){ try { cache.put(key, JSON.stringify(out), 21600); } catch (e){} }
  return out;
}

/* ===================== SHEET HELPERS ===================== */
function ss_(){
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Run setup() first');
  return SpreadsheetApp.openById(id);
}
const cell_ = v => v instanceof Date ? Utilities.formatDate(v, TZ, TS_FMT) : String(v == null ? '' : v).trim();
function tbl_(name){
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error(`Sheet "${name}" missing – run setup()`);
  const v = sh.getDataRange().getValues();
  const h = v[0].map(x => String(x).trim());
  const rows = v.slice(1).map((r, i) => { const o = { _row:i + 2 }; h.forEach((k, j) => o[k] = cell_(r[j])); return o; })
    .filter(o => h.some(k => o[k] !== ''));
  return { sh, h, rows };
}
function append_(t, obj){
  const row = t.h.map(k => obj[k] == null ? '' : obj[k]);
  t.sh.appendRow(row);
  const o = Object.assign({ _row:t.sh.getLastRow() }, obj);
  t.rows.push(o); return o;
}
function update_(t, row, ch){
  Object.assign(row, ch);
  t.sh.getRange(row._row, 1, 1, t.h.length).setValues([t.h.map(k => row[k] == null ? '' : row[k])]);
}
const ts_ = d => Utilities.formatDate(d || new Date(), TZ, TS_FMT);
function ms_(s){ if (!s) return 0; try { return Utilities.parseDate(s, TZ, TS_FMT).getTime(); } catch (e){ const t = Date.parse(s); return isNaN(t) ? 0 : t; } }
const yes_ = v => v === 'Y' || v === 'TRUE' || v === 'true';
const nrm_ = s => String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
const uid_ = () => Utilities.getUuid().replace(/-/g, '');
function hist_(by, role, table, key, field, oldV, newV, note){
  const sh = ss_().getSheetByName('History');
  sh.appendRow([ts_(), by, role, table, key, field, oldV == null ? '' : String(oldV), newV == null ? '' : String(newV), note || '']);
}
function addPoints_(username, delta, reason, ref, by){
  if (!delta) return;
  ss_().getSheetByName('Points').appendRow([ts_(), username, delta, reason, ref || '', by || 'system']);
  const t = tbl_('Users'), u = t.rows.find(r => r.Username === username);
  if (u) update_(t, u, { Points:(Number(u.Points) || 0) + delta });
}

/* ===================== AUTH ===================== */
function hash_(salt, pw){
  let h = salt + ':' + pw;
  for (let i = 0; i < 300; i++) h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
  return h;
}
function findUser_(username){ return tbl_('Users').rows.find(r => r.Username === username); }
function appUrl_(){ return ScriptApp.getService().getUrl(); }
function entryUrl_(){ return ENTRY_URL || appUrl_(); }
function mail_(to, subject, html){
  MailApp.sendEmail({ to, subject, htmlBody:html, name:'Radar Quest' });
}
function mailBox_(title, text, btn, url){
  return `<div style="font-family:Arial,sans-serif;max-width:480px;margin:auto;padding:24px;border-radius:16px;background:#EEF4FF;color:#1D2550">
    <h2 style="margin:0 0 8px">${title}</h2><p style="margin:0 0 18px">${text}</p>
    <a href="${url}" style="display:inline-block;background:#FFC23D;color:#24306E;font-weight:bold;text-decoration:none;padding:12px 22px;border-radius:12px">${btn}</a>
    <p style="font-size:12px;color:#5E6894;margin-top:18px">If the button doesn't work or you see "Page not found", copy this link and open it in Chrome:<br>${url}</p></div>`;
}
function register_(p){
  const email = String(p.email || '').trim().toLowerCase();
  const username = String(p.username || '').trim().toLowerCase();
  if (!EMAIL_RE.test(email)) return { error:'Check your email', field:'email' };
  if (!USER_RE.test(username)) return { error:'3–20 letters, numbers, . or _', field:'username' };
  if (String(p.pw || '').length < 6) return { error:'6+ characters', field:'pw' };
  const t = tbl_('Users');
  if (t.rows.find(r => r.Username === username)) return { error:'Username taken', field:'username' };
  const old = t.rows.find(r => r.Email === email);
  if (old && old.Status !== 'unverified') return { error:'Email already registered', field:'email' };
  const salt = Utilities.getUuid(), tok = uid_();
  const row = { Username:username, Email:email, Avatar:String(p.avatar || 'a1').slice(0, 10), 'Avatar Img':String(p.avatarImg || '').length < 40000 ? (p.avatarImg || '') : '',
    Role:'player', Status:'unverified', Salt:salt, PwHash:hash_(salt, p.pw), 'Act Token':tok, Created:ts_(), Points:0 };
  if (old) update_(t, old, row); else append_(t, row);
  const url = appUrl_() + '?act=' + tok;
  mail_(email, 'Radar Quest · Activate your account', mailBox_('Hi ' + username + '!', 'Tap the button to activate your Radar Quest account. After that, an admin will approve you.', 'Activate account', url));
  return { ok:true, email, username };
}
function resend_(p){
  const t = tbl_('Users'), u = t.rows.find(r => r.Email === String(p.email || '').trim().toLowerCase());
  if (!u || u.Status !== 'unverified') return { ok:true };
  const tok = uid_(); update_(t, u, { 'Act Token':tok });
  mail_(u.Email, 'Radar Quest · Activate your account', mailBox_('Hi ' + u.Username + '!', 'Tap the button to activate your Radar Quest account.', 'Activate account', appUrl_() + '?act=' + tok));
  return { ok:true };
}
function activate_(tok){
  return withLock_(() => {
    const t = tbl_('Users'), u = t.rows.find(r => r['Act Token'] && r['Act Token'] === tok);
    if (!u) return 'BAD';
    // the token stays valid, so opening the link twice (or a mail scanner opening it first) still shows "activated"
    if (u.Status === 'unverified'){ update_(t, u, { Status:'pending', 'Activated At':ts_() }); return 'OK'; }
    if (u.Status === 'pending') return 'OK';
    if (u.Status === 'active') return 'ACTIVE';
    return 'BAD';
  });
}
function login_(p){
  const email = String(p.email || '').trim().toLowerCase();
  const u = tbl_('Users').rows.find(r => r.Email === email || r.Username === email);
  if (!u || u.PwHash !== hash_(u.Salt, String(p.pw || ''))) return { error:'Wrong email or password' };
  if (u.Status === 'unverified') return { error:'UNVERIFIED', email:u.Email };
  if (u.Status === 'pending') return { error:'PENDING' };
  if (u.Status !== 'active') return { error:'Account not active · contact admin' };
  const token = uid_() + uid_().slice(0, 8);
  ss_().getSheetByName('Sessions').appendRow([token, u.Username, ts_(new Date(Date.now() + SESSION_DAYS * DAY))]);
  return Object.assign({ token }, bootstrap_(u));
}
function auth_(token){
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  let name = cache.get('s:' + token);
  if (!name){
    const row = tbl_('Sessions').rows.find(r => r.Token === token);
    if (!row || ms_(row.Expires) < Date.now()) return null;
    name = row.Username; cache.put('s:' + token, name, 600);
  }
  const u = findUser_(name);
  return u && u.Status === 'active' ? u : null;
}
function logout_(token){
  CacheService.getScriptCache().remove('s:' + token);
  const t = tbl_('Sessions'), r = t.rows.find(x => x.Token === token);
  if (r) t.sh.deleteRow(r._row);
  return { ok:true };
}
function forgot_(p){
  const email = String(p.email || '').trim().toLowerCase();
  return withLock_(() => {
    const t = tbl_('Users'), u = t.rows.find(r => r.Email === email);
    if (!u) return { error:'NOT_REGISTERED' };
    const tok = uid_();
    update_(t, u, { 'Reset Token':tok, 'Reset Expires':ts_(new Date(Date.now() + DAY)) });
    mail_(email, 'Radar Quest · Reset password', mailBox_('Reset your password', 'Tap the button to set a new password. The link works for 24 hours.', 'Set new password', appUrl_() + '?reset=' + tok));
    return { ok:true };
  });
}
function resetTokenOk_(tok){ const u = tbl_('Users').rows.find(r => r['Reset Token'] === tok); return !!(u && ms_(u['Reset Expires']) > Date.now()); }
function setNewPw_(p){
  const t = tbl_('Users'), u = t.rows.find(r => r['Reset Token'] && r['Reset Token'] === p.token);
  if (!u || ms_(u['Reset Expires']) < Date.now()) return { error:'Link expired · try again' };
  if (String(p.pw || '').length < 6) return { error:'6+ characters' };
  const salt = Utilities.getUuid();
  update_(t, u, { Salt:salt, PwHash:hash_(salt, p.pw), 'Reset Token':'', 'Reset Expires':'' });
  return { ok:true };
}
function changePw_(me, p){
  if (me.PwHash !== hash_(me.Salt, String(p.old || ''))) return { error:'Current password is wrong' };
  if (String(p.pw || '').length < 6) return { error:'6+ characters' };
  const t = tbl_('Users'), u = t.rows.find(r => r.Username === me.Username), salt = Utilities.getUuid();
  update_(t, u, { Salt:salt, PwHash:hash_(salt, p.pw) });
  return { ok:true };
}
function setAvatar_(me, p){
  const t = tbl_('Users'), u = t.rows.find(r => r.Username === me.Username);
  update_(t, u, { Avatar:String(p.avatar || 'a1').slice(0, 10), 'Avatar Img':String(p.avatarImg || '').length < 40000 ? (p.avatarImg || '') : '' });
  return { ok:true };
}

/* ===================== DATA ===================== */
function locations_(){
  return tbl_('Locations').rows.filter(r => r.Location && r.Active !== 'N').map(r => ({ name:r.Location, isNew:yes_(r.New), by:r['Added By'] }));
}
function userOut_(u){ return { username:u.Username, email:u.Email, avatar:u.Avatar, avatarImg:u['Avatar Img'], role:u.Role, points:Number(u.Points) || 0 }; }
function leaderboard_(users){
  return users.filter(r => r.Status === 'active' && r.Role !== 'admin').map(r => ({ username:r.Username, avatar:r.Avatar, avatarImg:r['Avatar Img'], points:Number(r.Points) || 0 }))
    .sort((a, b) => b.points - a.points);
}
function bootstrap_(u){
  const users = tbl_('Users').rows;
  const me = users.find(r => r.Username === u.Username) || u;
  const out = { user:userOut_(me), locations:locations_(), leaderboard:leaderboard_(users), statuses:STATUSES, pts:{ move:PTS_MOVE, big:PTS_BIG } };
  if (me.Role === 'admin') out.admin = adminCounts_(users);
  return out;
}
function adminCounts_(users){
  users = users || tbl_('Users').rows;
  return { pendingUsers:users.filter(r => r.Status === 'pending').length,
           pendingItems:tbl_('NewItems').rows.filter(r => r.State === 'Pending').length,
           toCheck:cols_('Updates', ['Flags', 'Check', 'State']).filter(r => r.Flags && !r.Check && r.State !== 'Reset').length,
           sheetUrl:ss_().getUrl() };
}

/* ---------- reference index (cached) ---------- */
function refIndex_(){
  const cache = CacheService.getScriptCache(), meta = cache.get('ref:n');
  if (meta){
    const keys = Array.from({ length:+meta }, (_, i) => 'ref:' + i), got = cache.getAll(keys);
    if (keys.every(k => got[k])) return JSON.parse(keys.map(k => got[k]).join(''));
  }
  const sh = ss_().getSheetByName('Ref'), last = sh.getLastRow();
  if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, 9).getValues();
  const idx = v.map((r, i) => [cell_(r[0]), cell_(r[1]), cell_(r[4]), cell_(r[3]).replace(/\s+/g, ' ').slice(0, 90), i + 2, yes_(cell_(r[8])) ? 1 : 0]);
  const s = JSON.stringify(idx), parts = [];
  for (let i = 0; i < s.length; i += 90000) parts.push(s.slice(i, i + 90000));
  const put = {}; parts.forEach((p, i) => put['ref:' + i] = p); put['ref:n'] = String(parts.length);
  try { cache.putAll(put, 1800); } catch (e){}
  return idx;
}
function clearRefCache_(){ try { CacheService.getScriptCache().remove('ref:n'); } catch (e){} }
function foundMap_(fresh){
  const cache = CacheService.getScriptCache();
  if (!fresh){ const c = cache.get('fm'); if (c) return JSON.parse(c); }
  const sh = ss_().getSheetByName('Updates'), last = sh.getLastRow(), map = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 13).getValues().forEach(r => {
    if (cell_(r[12]) === 'Reset') return;
    map[cell_(r[3])] = { id:cell_(r[0]), at:cell_(r[1]), by:cell_(r[2]), loc:cell_(r[7]), big:yes_(cell_(r[8])) };
  });
  try { cache.put('fm', JSON.stringify(map), 600); } catch (e){}   // too big for the cache → just read the sheet each time
  return map;
}
function clearFoundCache_(){ try { CacheService.getScriptCache().remove('fm'); } catch (e){} }
function readRef_(row){
  const sh = ss_().getSheetByName('Ref'), h = HEADERS.Ref, v = sh.getRange(row, 1, 1, h.length).getValues()[0];
  const o = { _row:row }; h.forEach((k, j) => o[k] = cell_(v[j])); return o;
}
const cleanDetail_ = d => String(d || '').replace(/\s+/g, ' ').replace(/\s*S\/N\s*\S+\s*$/i, '').trim();
function refOut_(o, found){
  return { key:o.Asset13, asset:o.Asset13, inv:o['Invent no'], serial:o.Serial, detail:cleanDetail_(o.Detail), detailFull:o.Detail,
    owner:o.Owner || o['Asset Sup'], status:o.Status, found:found || null };
}
function search_(p){
  const q = String(p.q || '').trim().toUpperCase(), qn = nrm_(q);
  if (q.length < 3) return { items:[], total:0 };
  const hits = refIndex_().filter(r => !r[5] && (nrm_(r[0]).includes(qn) || (qn && nrm_(r[1]).includes(qn)) || (qn && nrm_(r[2]).includes(qn)) || r[3].toUpperCase().includes(q)));
  const rank = r => [r[0], r[1], r[2]].some(x => nrm_(x) === qn) ? 0 : [r[0], r[1], r[2]].some(x => nrm_(x).startsWith(qn)) ? 1 : 2;
  hits.sort((a, b) => rank(a) - rank(b));
  const fm = foundMap_();
  return { items:hits.slice(0, 12).map(r => ({ key:r[0], asset:r[0], inv:r[1], serial:r[2], detail:cleanDetail_(r[3]), found:fm[r[0]] || null })), total:hits.length,
           news:newMatches_(qn) };
}
function newMatches_(qn){
  if (!qn || qn.length < 4) return [];
  return tbl_('NewItems').rows.filter(r => r.State !== 'Rejected' && nrm_(r.Number).includes(qn)).slice(0, 3)
    .map(r => ({ id:r.ID, number:r.Number, detail:r.Detail, by:r.Username, at:r.Timestamp, state:r.State, loc:r.Location }));
}
/** exact code(s) → one item (Ref) or an already reported new item */
function lookup_(p){
  const codes = (p.codes || [p.code]).map(nrm_).filter(c => c.length >= 4);
  if (!codes.length) return { item:null };
  const idx = refIndex_();
  for (const c of codes){
    const r = idx.find(x => !x[5] && (nrm_(x[0]) === c || nrm_(x[1]) === c || nrm_(x[2]) === c));
    if (r) return { item:refOut_(readRef_(r[4]), foundMap_()[r[0]]), code:c };
  }
  const n = tbl_('NewItems').rows.find(r => r.State !== 'Rejected' && codes.includes(nrm_(r.Number)));
  if (n) return { item:null, reported:{ id:n.ID, number:n.Number, detail:n.Detail, by:n.Username, at:n.Timestamp, state:n.State, loc:n.Location } };
  return { item:null };
}

/* ---------- OCR ---------- */
function blob_(dataUrl, name){
  const m = /^data:([^;]+);base64,(.*)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('Bad image');
  return Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
}
function ocr_(blob){
  let id;
  if (Drive.Files.insert) id = Drive.Files.insert({ title:'rq-ocr', mimeType:blob.getContentType() }, blob, { ocr:true, ocrLanguage:'en' }).id;
  else id = Drive.Files.create({ name:'rq-ocr', mimeType:MimeType.GOOGLE_DOCS }, blob, { ocrLanguage:'en', fields:'id' }).id;
  try { return DocumentApp.openById(id).getBody().getText(); }
  finally { try { DriveApp.getFileById(id).setTrashed(true); } catch (e){} }
}
const fuzz_ = s => nrm_(s).replace(/[OQ]/g, '0').replace(/[IL]/g, '1').replace(/B/g, '8').replace(/S/g, '5').replace(/Z/g, '2');
function tokens_(text){
  const raw = String(text || '').replace(/S\s*\/\s*N\s*[:.]?/gi, ' ').split(/\s+/).map(t => t.replace(/^[^A-Za-z0-9]+|[^A-Za-z0-9]+$/g, '')).filter(Boolean);
  const out = raw.slice(); for (let i = 0; i < raw.length - 1; i++) out.push(raw[i] + raw[i + 1]);
  return out.filter(t => nrm_(t).length >= 5 && /\d/.test(t));
}
function identify_(p){
  let text;
  try { text = ocr_(blob_(p.image, 'label.jpg')); } catch (e){ return { item:null, tokens:[], error:'OCR unavailable' }; }
  const toks = tokens_(text);
  const idx = refIndex_();
  for (const t of toks){
    const f = fuzz_(t);
    const r = idx.find(x => !x[5] && [x[0], x[1], x[2]].some(c => c && fuzz_(c) === f));
    if (r) return { item:refOut_(readRef_(r[4]), foundMap_()[r[0]]), read:t, tokens:toks.slice(0, 6) };
  }
  const look = lookup_({ codes:toks });
  return { item:null, reported:look.reported || null, tokens:toks.sort((a, b) => b.length - a.length).slice(0, 6) };
}
/** Read the label photo and compare with the chosen item (used when the player typed the number). */
function checkLabel_(dataUrl, it){
  const targets = [it.asset, it.serial, it.inv].filter(t => nrm_(t).length >= 4);
  if (!targets.length) return { state:'na', read:'' };
  let text; try { text = ocr_(blob_(dataUrl, 'label.jpg')); } catch (e){ return { state:'unavailable', read:'' }; }
  const toks = tokens_(text);
  if (!toks.length) return { state:'unreadable', read:'' };
  for (const t of targets){ const f = fuzz_(t); const hit = toks.find(x => fuzz_(x).includes(f) || f.includes(fuzz_(x)) && fuzz_(x).length >= 8); if (hit) return { state:'match', read:hit }; }
  return { state:'mismatch', read:toks.sort((a, b) => b.length - a.length)[0].toUpperCase().slice(0, 40) };
}
/** Photo fingerprints (256-bit dHash, hex) → ids of earlier finds that used (almost) the same photo. */
function hamming_(a, b){
  if (!a || !b || a.length !== b.length) return 999;
  let d = 0; for (let i = 0; i < a.length; i++){ let x = parseInt(a[i], 16) ^ parseInt(b[i], 16); while (x){ d += x & 1; x >>= 1; } }
  return d;
}
function dupPhotos_(hashes){
  hashes = (hashes || []).map(h => String(h || '').toLowerCase()).map(h => /^[0-9a-f]{64}$/.test(h) ? h : '');
  if (!hashes.some(Boolean)) return [];
  const out = {};
  cols_('Updates', ['ID', 'Photo Hash']).forEach(r => {
    const old = String(r['Photo Hash'] || '').split('|');
    hashes.forEach((h, i) => { if (!h) return; old.forEach((o, j) => { if (o && hamming_(h, o) <= (i === 2 || j === 2 ? 6 : 10)) out[r.ID] = 1; }); });
  });
  return Object.keys(out).slice(0, 3);
}

/* ---------- photos ---------- */
function folder_(key, name){
  const props = PropertiesService.getScriptProperties(), id = props.getProperty(key);
  if (id){ try { return DriveApp.getFolderById(id); } catch (e){} }
  const f = DriveApp.createFolder(name); props.setProperty(key, f.getId()); return f;
}
function savePhoto_(dataUrl, name){ return dataUrl ? folder_('PHOTO_FOLDER_ID', 'Radar Quest Photos').createFile(blob_(dataUrl, name)).getUrl() : ''; }
const thumbOk_ = s => String(s || '').length < 45000 ? (s || '') : '';
const fileId_ = url => (/\/d\/([\w-]+)/.exec(url || '') || /id=([\w-]+)/.exec(url || '') || [])[1] || '';

/* ---------- locations ---------- */
function ensureLocation_(me, name){
  name = String(name || '').replace(/\s+/g, ' ').trim().slice(0, 120);
  if (!name) throw new Error('Pick a location');
  const t = tbl_('Locations'), hit = t.rows.find(r => r.Location.toLowerCase() === name.toLowerCase());
  if (hit) return hit.Location;
  append_(t, { Location:name, New:'Y', 'Added By':me.Username, 'Added At':ts_(), Active:'Y' });
  hist_(me.Username, me.Role, 'Locations', name, 'Location', '', name, 'added');
  return name;
}
function addLocation_(me, p){ const name = ensureLocation_(me, p.name); return { name, locations:locations_() }; }
function editLocation_(me, p){
  const t = tbl_('Locations'), r = t.rows.find(x => x.Location === p.old);
  if (!r) return { error:'Not found' };
  const ch = {};
  if (p.name != null && p.name.trim() && p.name.trim() !== r.Location){ ch.Location = p.name.trim(); hist_(me.Username, 'admin', 'Locations', r.Location, 'Location', r.Location, ch.Location); }
  if (p.clearNew){ ch.New = ''; }
  if (p.hide){ ch.Active = 'N'; hist_(me.Username, 'admin', 'Locations', r.Location, 'Active', 'Y', 'N'); }
  update_(t, r, ch);
  return { locations:locations_() };
}

/* ===================== PLAYER WRITES ===================== */
const trashUrls_ = urls => urls.forEach(u => { try { if (fileId_(u)) DriveApp.getFileById(fileId_(u)).setTrashed(true); } catch (e){} });
function submitFind_(me, d){
  const look = lookup_({ codes:[d.key] });
  if (!look.item) return { error:'Item not found' };
  const it = look.item;
  if (it.found) return { error:'ALREADY', found:it.found };
  if (!d.photo1 || !d.photo2 || !d.photo3) return { error:'3 photos required' };
  // photos are saved outside the lock so several players can submit at the same time
  const tag = (it.serial || it.asset).replace(/[^\w.-]+/g, '_') + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss');
  const urls = [savePhoto_(d.photo1, tag + '_1-barcode.jpg'), savePhoto_(d.photo2, tag + '_2-device.jpg'), savePhoto_(d.photo3, tag + '_3-wide.jpg')];
  // anti-cheat flags for the admin "Needs check" list (points are still given right away)
  const pick = d.pick === 'photo' ? 'photo' : 'typed';
  const chk = pick === 'photo' ? { state:'match', read:String(d.read || '') } : checkLabel_(d.photo1, it);
  const age = Math.max(0, Math.round(Number(d.photoAge) || 0));
  const dups = dupPhotos_(d.hashes);
  const flags = [];
  if (pick === 'typed' && ['mismatch', 'unreadable', 'unavailable'].includes(chk.state)) flags.push('label-' + chk.state);
  if (age > 15) flags.push('old-photo');
  dups.forEach(id => flags.push('dup:' + id));
  const out = withLock_(() => {
    const again = foundMap_(true)[it.asset];
    if (again) return { error:'ALREADY', found:again };
    const loc = ensureLocation_(me, d.location);
    const big = !!d.tooBig, pts = big ? PTS_BIG : PTS_MOVE;
    const id = 'F' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    const row = { ID:id, Timestamp:ts_(), Username:me.Username, Asset13:it.asset, Serial:it.serial, 'Invent no':it.inv, Detail:it.detail, Location:loc,
      'Too Big':big ? 'Y' : '', 'Barcode Check':chk.state, 'Barcode Read':String(chk.read || '').slice(0, 60), Points:pts, State:'Active',
      'Photo Barcode':urls[0], 'Photo Device':urls[1], 'Photo Wide':urls[2], Thumb1:thumbOk_(d.thumb1), Thumb2:thumbOk_(d.thumb2), Thumb3:thumbOk_(d.thumb3),
      Pick:pick, 'Photo Age':String(age), 'Photo Hash':(d.hashes || []).map(h => /^[0-9a-f]{64}$/i.test(String(h || '')) ? String(h).toLowerCase() : '').join('|'), Flags:flags.join(',') };
    append_(tbl_('Updates'), row);
    clearFoundCache_();
    hist_(me.Username, me.Role, 'Updates', it.asset, 'Found', '', loc, big ? 'too big to move' : 'moved to ASSET ZONE');
    addPoints_(me.Username, pts, big ? 'Find (too big to move)' : 'Find', id, 'system');
    const u = findUser_(me.Username);
    return { id, pts, points:Number(u.Points) || 0, leaderboard:leaderboard_(tbl_('Users').rows), locations:locations_() };
  });
  if (out.error) trashUrls_(urls);
  return out;
}
function submitNew_(me, d){
  const number = String(d.number || '').trim().toUpperCase().slice(0, 60);
  if (nrm_(number).length < 4) return { error:'Asset or Serial No. needed' };
  const look = lookup_({ codes:[number] });
  if (look.item) return { error:'IN_LIST', item:look.item };
  if (look.reported) return { error:'REPORTED', reported:look.reported };
  if (!String(d.detail || '').trim()) return { error:'Add a detail' };
  if (!d.photo1 || !d.photo2 || !d.photo3) return { error:'3 photos required' };
  const tag = 'NEW_' + nrm_(number) + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss');
  const urls = [savePhoto_(d.photo1, tag + '_1-barcode.jpg'), savePhoto_(d.photo2, tag + '_2-device.jpg'), savePhoto_(d.photo3, tag + '_3-wide.jpg')];
  const out = withLock_(() => {
    const t = tbl_('NewItems');
    const dup = t.rows.find(r => r.State !== 'Rejected' && nrm_(r.Number) === nrm_(number));
    if (dup) return { error:'REPORTED', reported:{ id:dup.ID, number:dup.Number, detail:dup.Detail, by:dup.Username, at:dup.Timestamp, state:dup.State, loc:dup.Location } };
    const loc = ensureLocation_(me, d.location);
    const n = t.rows.reduce((m, r) => Math.max(m, Number(String(r.ID).replace(/\D/g, '')) || 0), 0) + 1;
    const id = 'NEW-' + ('000' + n).slice(-4);
    append_(t, { ID:id, Timestamp:ts_(), Username:me.Username, Number:number, Detail:String(d.detail).trim().slice(0, 200), Note:String(d.note || '').trim().slice(0, 300),
      Location:loc, 'Too Big':d.tooBig ? 'Y' : '', 'Barcode Read':String(d.read || '').slice(0, 60), State:'Pending', Points:0,
      'Photo Barcode':urls[0], 'Photo Device':urls[1], 'Photo Wide':urls[2], Thumb1:thumbOk_(d.thumb1), Thumb2:thumbOk_(d.thumb2), Thumb3:thumbOk_(d.thumb3) });
    hist_(me.Username, me.Role, 'NewItems', id, 'Created', '', number, String(d.detail).trim());
    return { id, pending:true, locations:locations_() };
  });
  if (out.error) trashUrls_(urls);
  return out;
}
function editNew_(me, p){
  const t = tbl_('NewItems'), r = t.rows.find(x => x.ID === p.id);
  if (!r) return { error:'Not found' };
  if (me.Role !== 'admin' && r.Username !== me.Username) return { error:'Only the finder or an admin can edit' };
  const fields = { number:'Number', detail:'Detail', note:'Note', location:'Location' }, ch = {};
  Object.keys(fields).forEach(k => {
    if (p[k] == null) return;
    let v = String(p[k]).trim();
    if (k === 'number') v = v.toUpperCase();
    if (k === 'location') v = ensureLocation_(me, v);
    if (v !== r[fields[k]]){ hist_(me.Username, me.Role, 'NewItems', r.ID, fields[k], r[fields[k]], v); ch[fields[k]] = v; }
  });
  if (Object.keys(ch).length) update_(t, r, ch);
  return { ok:true, item:newOut_(r) };
}
function newOut_(r){
  return { id:r.ID, at:r.Timestamp, by:r.Username, number:r.Number, detail:r.Detail, note:r.Note, location:r.Location, tooBig:yes_(r['Too Big']), state:r.State,
    division:r.Division, remark:r.Remark, status:r.Status, owner:r.Owner, points:Number(r.Points) || 0, reviewedBy:r['Reviewed By'],
    thumbs:[r.Thumb1, r.Thumb2, r.Thumb3], photos:[r['Photo Barcode'], r['Photo Device'], r['Photo Wide']] };
}
function findOut_(r){
  return { id:r.ID, at:r.Timestamp, by:r.Username, key:r.Asset13, serial:r.Serial, inv:r['Invent no'], detail:r.Detail, location:r.Location, tooBig:yes_(r['Too Big']),
    check:r['Barcode Check'], read:r['Barcode Read'], pick:r.Pick, age:r['Photo Age'], flags:r.Flags ? String(r.Flags).split(',') : [], verdict:r.Check, checkedBy:r['Checked By'],
    points:Number(r.Points) || 0, state:r.State, thumbs:[r.Thumb1, r.Thumb2, r.Thumb3], photos:[r['Photo Barcode'], r['Photo Device'], r['Photo Wide']] };
}
function myFinds_(me){
  const f = tbl_('Updates').rows.filter(r => r.Username === me.Username).slice(-60).reverse().map(findOut_);
  const n = tbl_('NewItems').rows.filter(r => r.Username === me.Username).slice(-60).reverse().map(newOut_);
  const pts = tbl_('Points').rows.filter(r => r.Username === me.Username && Number(r.Delta) < 0).slice(-20).reverse()
    .map(r => ({ at:r.Timestamp, delta:Number(r.Delta), reason:r.Reason, ref:r.Ref }));
  return { finds:f, news:n, deductions:pts };
}
/** One full-size photo of a find / new item: only its finder or an admin. */
function photo_(me, p){
  const i = Number(p.i), cols = ['Photo Barcode', 'Photo Device', 'Photo Wide'];
  if (!(i >= 0 && i < 3)) return { error:'Bad photo' };
  const id = String(p.id || ''), t = tbl_(/^NEW-/.test(id) ? 'NewItems' : 'Updates'), r = t.rows.find(x => x.ID === id);
  if (!r) return { error:'Not found' };
  if (me.Role !== 'admin' && r.Username !== me.Username) return { error:'Not yours' };
  const fid = fileId_(r[cols[i]]);
  if (!fid) return { error:'No photo' };
  const b = DriveApp.getFileById(fid).getBlob();
  return { src:'data:' + (b.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(b.getBytes()) };
}

/* ===================== ADMIN ===================== */
function adminHome_(){
  const users = tbl_('Users').rows;
  return Object.assign(adminCounts_(users), {
    pendingList:tbl_('NewItems').rows.filter(r => r.State === 'Pending').map(newOut_),
    recentNew:tbl_('NewItems').rows.filter(r => r.State !== 'Pending').slice(-20).reverse().map(newOut_),
    locations:locations_() });
}
function members_(){
  return { members:tbl_('Users').rows.map(r => ({ username:r.Username, email:r.Email, avatar:r.Avatar, avatarImg:r['Avatar Img'], role:r.Role, status:r.Status, created:r.Created, points:Number(r.Points) || 0 })) };
}
function setUser_(me, p){
  const t = tbl_('Users'), u = t.rows.find(r => r.Username === p.username);
  if (!u) return { error:'Not found' };
  const ch = {};
  if (p.status && ['active', 'rejected', 'suspended', 'pending'].includes(p.status) && p.status !== u.Status){
    if (u.Username === me.Username) return { error:"Can't change yourself" };
    hist_(me.Username, 'admin', 'Users', u.Username, 'Status', u.Status, p.status); ch.Status = p.status; ch['Reviewed By'] = me.Username; ch['Reviewed At'] = ts_();
  }
  if (p.role && ['player', 'admin'].includes(p.role) && p.role !== u.Role){
    if (u.Username === me.Username) return { error:"Can't change yourself" };
    hist_(me.Username, 'admin', 'Users', u.Username, 'Role', u.Role, p.role); ch.Role = p.role;
  }
  const approved = ch.Status === 'active' && u.Status !== 'active' && u.Status !== 'suspended';
  if (Object.keys(ch).length) update_(t, u, ch);
  if (approved){
    try { mail_(u.Email, 'Radar Quest · You are in!', mailBox_('Welcome, ' + u.Username + '!', 'An admin approved your account. Log in with your email and password and start finding gear.', 'Open Radar Quest', entryUrl_())); }
    catch (e){ console.error('approval mail', e); }
  }
  return members_();
}
function reviewNew_(me, p){
  const t = tbl_('NewItems'), r = t.rows.find(x => x.ID === p.id);
  if (!r) return { error:'Not found' };
  const ch = {};
  ['division:Division', 'remark:Remark', 'status:Status', 'owner:Owner', 'detail:Detail', 'number:Number'].forEach(s => {
    const [k, f] = s.split(':');
    if (p[k] != null && String(p[k]).trim() !== r[f]){ hist_(me.Username, 'admin', 'NewItems', r.ID, f, r[f], String(p[k]).trim()); ch[f] = String(p[k]).trim(); }
  });
  if (p.decision && ['Approved', 'Rejected', 'Pending'].includes(p.decision) && p.decision !== r.State){
    hist_(me.Username, 'admin', 'NewItems', r.ID, 'State', r.State, p.decision);
    ch.State = p.decision; ch['Reviewed By'] = me.Username; ch['Reviewed At'] = ts_();
    const want = p.decision === 'Approved' ? (yes_(r['Too Big']) ? PTS_BIG : PTS_MOVE) : 0, had = Number(r.Points) || 0;
    if (want !== had){ addPoints_(r.Username, want - had, p.decision === 'Approved' ? 'New item approved' : 'New item ' + p.decision.toLowerCase(), r.ID, me.Username); ch.Points = want; }
  }
  if (Object.keys(ch).length) update_(t, r, ch);
  return { item:newOut_(r), counts:adminCounts_() };
}
function finds_(p){
  const q = nrm_(p.q || ''), ql = String(p.q || '').toLowerCase();
  let rows = tbl_('Updates').rows.slice().reverse();
  if (q || ql) rows = rows.filter(r => nrm_(r.Asset13).includes(q) || nrm_(r.Serial).includes(q) || r.Username.includes(ql) || r.Location.toLowerCase().includes(ql) || r.Detail.toLowerCase().includes(ql));
  return { finds:rows.slice(0, 80).map(findOut_), total:rows.length };
}
function checks_(){
  const rows = tbl_('Updates').rows.filter(r => r.Flags && !r.Check && r.State !== 'Reset').reverse();
  return { checks:rows.slice(0, 50).map(findOut_), total:rows.length };
}
function setCheck_(me, p){
  const t = tbl_('Updates'), r = t.rows.find(x => x.ID === p.id);
  if (!r) return { error:'Not found' };
  const verdict = p.verdict === 'Wrong' ? 'Wrong' : 'OK';
  hist_(me.Username, 'admin', 'Updates', r.ID, 'Check', r.Check, verdict, String(p.reason || ''));
  update_(t, r, { Check:verdict, 'Checked By':me.Username, 'Checked At':ts_() });
  const n = Math.abs(Math.round(Number(p.deduct) || 0));
  if (verdict === 'Wrong' && n) addPoints_(r.Username, -n, String(p.reason || 'Wrong item').slice(0, 200), r.ID, me.Username);
  return Object.assign(checks_(), { counts:adminCounts_() });
}
function assetAdmin_(p){
  const look = lookup_({ codes:[p.key] });
  if (!look.item) return { error:'Not found' };
  const finds = tbl_('Updates').rows.filter(r => r.Asset13 === look.item.asset).reverse().map(findOut_);
  const h = tbl_('History').rows.filter(r => r.Key === look.item.asset || finds.some(f => f.id === r.Key)).slice(-30).reverse()
    .map(r => ({ at:r.Timestamp, by:r.By, field:r.Field, old:r.Old, nw:r.New, note:r.Note }));
  const pts = tbl_('Points').rows.filter(r => finds.some(f => f.id === r.Ref)).map(r => ({ at:r.Timestamp, user:r.Username, delta:Number(r.Delta), reason:r.Reason, by:r.By, ref:r.Ref }));
  return { item:look.item, finds, history:h, points:pts, statuses:STATUSES };
}
function editAsset_(me, p){
  const look = lookup_({ codes:[p.key] });
  if (!look.item) return { error:'Not found' };
  const sh = ss_().getSheetByName('Ref'), idx = refIndex_().find(x => x[0] === look.item.asset);
  const r = readRef_(idx[4]), ch = {};
  if (p.status != null && p.status !== r.Status){ if (p.status && !STATUSES.includes(p.status)) return { error:'Bad status' }; ch.Status = p.status; }
  if (p.owner != null && String(p.owner).trim() !== (r.Owner || r['Asset Sup'])) ch.Owner = String(p.owner).trim();
  if (p.note != null && String(p.note).trim() !== r['Admin Note']) ch['Admin Note'] = String(p.note).trim();
  if (p.removed != null){ const v = p.removed ? 'Y' : ''; if (v !== r.Removed) ch.Removed = v; }
  Object.keys(ch).forEach(k => hist_(me.Username, 'admin', 'Ref', r.Asset13, k, k === 'Owner' ? (r.Owner || r['Asset Sup']) : r[k], ch[k]));
  if (Object.keys(ch).length){
    ch['Updated By'] = me.Username; ch['Updated At'] = ts_();
    Object.assign(r, ch);
    sh.getRange(r._row, 1, 1, HEADERS.Ref.length).setValues([HEADERS.Ref.map(k => r[k] == null ? '' : r[k])]);
    clearRefCache_();
  }
  if (p.findId && p.location != null){
    const t = tbl_('Updates'), f = t.rows.find(x => x.ID === p.findId), loc = ensureLocation_(me, p.location);
    if (f && f.Location !== loc){ hist_(me.Username, 'admin', 'Updates', f.ID, 'Location', f.Location, loc); update_(t, f, { Location:loc }); }
  }
  return assetAdmin_({ key:r.Asset13 });
}
function addAsset_(me, p){
  const a = String(p.asset || '').trim();
  if (!/^\d{6,20}$/.test(a)) return { error:'Asset13 needed (numbers)' };
  if (refIndex_().some(x => x[0] === a)) return { error:'Already in list' };
  const row = { Asset13:a, 'Invent no':String(p.inv || '').trim(), 'Asset Sup':String(p.owner || '').trim(), Detail:String(p.detail || '').trim(), Serial:String(p.serial || '').trim().toUpperCase(),
    'Added By':me.Username, 'Updated By':me.Username, 'Updated At':ts_() };
  const sh = ss_().getSheetByName('Ref');
  sh.appendRow(HEADERS.Ref.map(k => row[k] == null ? '' : row[k]));
  sh.getRange(sh.getLastRow(), 1, 1, HEADERS.Ref.length).setNumberFormat('@');
  hist_(me.Username, 'admin', 'Ref', a, 'Inserted', '', row.Detail, p.fromNew || '');
  clearRefCache_();
  if (p.fromNew){
    const t = tbl_('NewItems'), r = t.rows.find(x => x.ID === p.fromNew);
    if (r) update_(t, r, { Remark:(r.Remark ? r.Remark + ' · ' : '') + 'Added to asset list as ' + a });
  }
  return { ok:true };
}
function adjustPoints_(me, p){
  const delta = Math.round(Number(p.delta) || 0);
  if (!delta) return { error:'Enter points' };
  if (!String(p.reason || '').trim()) return { error:'Add a reason' };
  if (!findUser_(p.username)) return { error:'User not found' };
  addPoints_(p.username, delta, String(p.reason).trim().slice(0, 200), p.ref || '', me.Username);
  hist_(me.Username, 'admin', 'Points', p.ref || p.username, 'Points', '', String(delta), String(p.reason).trim());
  return { ok:true, points:Number(findUser_(p.username).Points) || 0 };
}
function resetFind_(me, p){
  const t = tbl_('Updates'), r = t.rows.find(x => x.ID === p.id);
  if (!r) return { error:'Not found' };
  hist_(me.Username, 'admin', 'Updates', r.ID, 'State', r.State, 'Reset', String(p.reason || ''));
  update_(t, r, { State:'Reset' });
  clearFoundCache_();
  if (p.removePoints && Number(r.Points)) addPoints_(r.Username, -Number(r.Points), 'Find reset: ' + String(p.reason || ''), r.ID, me.Username);
  return assetAdmin_({ key:r.Asset13 });
}

/* ===================== EXPORT (zip: Excel + photos) ===================== */
function exportZip_(p){
  const from = p.from || '0000', to = (p.to || '9999') + ' 99';
  const inRange = r => r.Timestamp >= from && r.Timestamp <= to;
  const finds = tbl_('Updates').rows.filter(inRange), news = tbl_('NewItems').rows.filter(inRange);
  if (!finds.length && !news.length) return { error:'No data in this range' };
  const refRows = {}; refIndex_().forEach(x => refRows[x[0]] = x[4]);
  const photoName = (prefix, n, kind) => `photos/${prefix}_${n}-${kind}.jpg`;
  const blobs = [], fCols = ['ID','Timestamp','Username','Asset13','Serial','Invent no','Detail','Location','Too Big','Points','State','Status (admin)','Owner (admin)','Admin Note','Photo 1','Photo 2','Photo 3'];
  const fData = [fCols];
  finds.forEach(r => {
    const ref = refRows[r.Asset13] ? readRef_(refRows[r.Asset13]) : {};
    const prefix = r.Asset13 + '_' + r.Timestamp.slice(0, 10);
    const names = ['barcode', 'device', 'wide'].map((k, i) => photoName(prefix, i + 1, k));
    [r['Photo Barcode'], r['Photo Device'], r['Photo Wide']].forEach((u, i) => { if (fileId_(u) && p.photos !== false) blobs.push([fileId_(u), names[i]]); });
    fData.push([r.ID, r.Timestamp, r.Username, r.Asset13, r.Serial, r['Invent no'], r.Detail, r.Location, r['Too Big'], r.Points, r.State, ref.Status || '', ref.Owner || ref['Asset Sup'] || '', ref['Admin Note'] || ''].concat(names));
  });
  const nCols = ['ID','Timestamp','Username','Number','Detail','Note','Location','Too Big','State','Division','Remark','Status','Owner','Points','Reviewed By','Photo 1','Photo 2','Photo 3'];
  const nData = [nCols];
  news.forEach(r => {
    const prefix = r.ID + '_' + nrm_(r.Number);
    const names = ['barcode', 'device', 'wide'].map((k, i) => photoName(prefix, i + 1, k));
    [r['Photo Barcode'], r['Photo Device'], r['Photo Wide']].forEach((u, i) => { if (fileId_(u) && p.photos !== false) blobs.push([fileId_(u), names[i]]); });
    nData.push([r.ID, r.Timestamp, r.Username, r.Number, r.Detail, r.Note, r.Location, r['Too Big'], r.State, r.Division, r.Remark, r.Status, r.Owner, r.Points, r['Reviewed By']].concat(names));
  });
  // build xlsx through a temporary spreadsheet
  const tmp = SpreadsheetApp.create('Radar Quest export tmp');
  const s1 = tmp.getSheets()[0].setName('Finds');
  s1.getRange(1, 1, fData.length, fCols.length).setNumberFormat('@').setValues(fData); s1.setFrozenRows(1);
  const s2 = tmp.insertSheet('New items');
  s2.getRange(1, 1, nData.length, nCols.length).setNumberFormat('@').setValues(nData); s2.setFrozenRows(1);
  [s1, s2].forEach(s => s.getRange(1, 1, 1, s.getLastColumn()).setFontWeight('bold').setBackground('#24306E').setFontColor('#FFFFFF'));
  SpreadsheetApp.flush();
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmm');
  const xlsx = UrlFetchApp.fetch(`https://docs.google.com/spreadsheets/d/${tmp.getId()}/export?format=xlsx`, { headers:{ Authorization:'Bearer ' + ScriptApp.getOAuthToken() } })
    .getBlob().setName(`radar-quest_${p.from || 'all'}_${p.to || 'all'}.xlsx`);
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  // pack, max ~40 MB per zip
  const out = folder_('EXPORT_FOLDER_ID', 'Radar Quest Exports'), parts = [];
  let cur = [xlsx], size = xlsx.getBytes().length;
  blobs.forEach(([id, name]) => {
    let b; try { b = DriveApp.getFileById(id).getBlob().setName(name); } catch (e){ return; }
    const n = b.getBytes().length;
    if (size + n > 40 * 1024 * 1024 && cur.length){ parts.push(cur); cur = []; size = 0; }
    cur.push(b); size += n;
  });
  parts.push(cur);
  const files = parts.map((list, i) => {
    const f = out.createFile(Utilities.zip(list, `radar-quest-export_${stamp}${parts.length > 1 ? '_part' + (i + 1) : ''}.zip`));
    return { name:f.getName(), url:f.getUrl(), download:'https://drive.google.com/uc?export=download&id=' + f.getId(), size:f.getSize() };
  });
  return { files, finds:finds.length, news:news.length, photos:blobs.length };
}

/* ===================== SETUP (run from the editor) ===================== */
/** Editor-only functions: public names can also be called from the web page, so refuse unless run by the owner. */
function ownerOnly_(){
  const me = String(Session.getActiveUser().getEmail() || '').toLowerCase(), owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!me || me !== owner) throw new Error('Run this from the Apps Script editor');
}
function setup(){
  ownerOnly_();
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('SHEET_ID');
  if (id){ try { ss = SpreadsheetApp.openById(id); } catch (e){} }
  if (!ss){ ss = SpreadsheetApp.create('Radar Quest Game Data'); props.setProperty('SHEET_ID', ss.getId()); }
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0){
      if (sh.getMaxColumns() < HEADERS[name].length) sh.insertColumnsAfter(sh.getMaxColumns(), HEADERS[name].length - sh.getMaxColumns());
      sh.getRange(1, 1, sh.getMaxRows(), HEADERS[name].length).setNumberFormat('@');
      sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold').setBackground('#24306E').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
    }
  });
  const def = ss.getSheetByName('Sheet1'); if (def && def.getLastRow() === 0) ss.deleteSheet(def);
  ss.getSheetByName('Sessions').hideSheet();
  // colour NewItems by state
  const ni = ss.getSheetByName('NewItems'), col = HEADERS.NewItems.indexOf('State') + 1;
  const letter = String.fromCharCode(64 + col), rng = ni.getRange(2, 1, ni.getMaxRows() - 1, HEADERS.NewItems.length);
  ni.setConditionalFormatRules([
    ['Pending', '#FFE3C7'], ['Approved', '#D9F5EC'], ['Rejected', '#E4E6EE'],
  ].map(([v, c]) => SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$${letter}2="${v}"`).setBackground(c).setRanges([rng]).build()));
  // starting locations
  const loc = tbl_('Locations');
  START_LOCATIONS.forEach(n => { if (!loc.rows.find(r => r.Location === n)) append_(loc, { Location:n, New:'', 'Added By':'setup', 'Added At':ts_(), Active:'Y' }); });
  // first admin = the script owner
  const users = tbl_('Users'), email = Session.getEffectiveUser().getEmail().toLowerCase();
  if (!users.rows.find(r => r.Role === 'admin')){
    const salt = Utilities.getUuid(), temp = 'rq' + Math.random().toString(36).slice(2, 8);
    append_(users, { Username:'admin', Email:email, Avatar:'a3', Role:'admin', Status:'active', Salt:salt, PwHash:hash_(salt, temp), Created:ts_(), Points:0 });
    console.log(`Admin login → email: ${email}  ·  temporary password: ${temp}  (change it in the app: Profile › Change password)`);
  }
  folder_('PHOTO_FOLDER_ID', 'Radar Quest Photos');
  schema_(true);
  console.log('Setup done. Sheet: ' + ss.getUrl());
}

const SAP_FILE = 'Asset_Master_SAP.xlsx';
const GIF_FILE = 'radar-quest-how-to-earn.gif';
function driveFile_(name){
  const it = DriveApp.getFilesByName(name);
  if (!it.hasNext()) throw new Error('Upload "' + name + '" to Google Drive first');
  return it.next();
}
function guideGif_(){
  // newest "radar-quest-how-to-earn.gif" in Drive (not in the trash); looked up again every hour so a new upload shows up
  const cache = CacheService.getScriptCache(); let id = cache.get('gif:id'), f = null;
  if (id){ try { f = DriveApp.getFileById(id); if (f.isTrashed()) f = null; } catch (e){ f = null; } }
  if (!f){
    const it = DriveApp.searchFiles('title = "' + GIF_FILE + '" and trashed = false');
    while (it.hasNext()){ const x = it.next(); if (!f || x.getLastUpdated() > f.getLastUpdated()) f = x; }
    if (!f) throw new Error('Upload "' + GIF_FILE + '" to Google Drive first');
    cache.put('gif:id', f.getId(), 3600);
  }
  return Utilities.base64Encode(f.getBlob().getBytes());
}
/** Load SAP Master_data from "Asset_Master_SAP.xlsx" in Google Drive (sheet Master_data). Safe to run again. */
function importSAP(){
  ownerOnly_();
  const src = driveFile_(SAP_FILE);
  let tmpId;
  if (Drive.Files.insert) tmpId = Drive.Files.insert({ title:'rq-sap-tmp', mimeType:MimeType.GOOGLE_SHEETS }, src.getBlob(), { convert:true }).id;
  else tmpId = Drive.Files.create({ name:'rq-sap-tmp', mimeType:MimeType.GOOGLE_SHEETS }, src.getBlob(), { fields:'id' }).id;
  let rows;
  try {
    const ws = SpreadsheetApp.openById(tmpId).getSheetByName('Master_data');
    if (!ws) throw new Error('Sheet "Master_data" not found in ' + SAP_FILE);
    rows = ws.getDataRange().getDisplayValues();
  } finally { try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e){} }
  const head = rows.shift().map(h => String(h).trim());
  const ix = k => head.indexOf(k);
  const need = ['Asset13', 'Invent no', 'Asset Sup', 'Detail', 'Serial'];
  need.forEach(k => { if (ix(k) < 0) throw new Error('Missing column ' + k); });
  const sh = ss_().getSheetByName('Ref');
  const keep = {};   // keep admin fields if re-importing
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.Ref.length).getValues().forEach(r => { keep[cell_(r[0])] = r.slice(5).map(cell_); });
  const seen = {}, data = [];
  rows.forEach(r => {
    const a = String(r[ix('Asset13')] || '').trim(); if (!a || seen[a]) return; seen[a] = 1;
    const base = need.map(k => String(r[ix(k)] == null ? '' : r[ix(k)]).trim());
    data.push(base.concat(keep[a] || ['', '', '', '', 'SAP import', '', '']));
  });
  if (sh.getMaxRows() < data.length + 1) sh.insertRowsAfter(sh.getMaxRows(), data.length + 1 - sh.getMaxRows());
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.Ref.length).clearContent();
  sh.getRange(2, 1, data.length, HEADERS.Ref.length).setNumberFormat('@');
  for (let i = 0; i < data.length; i += 2000){ const part = data.slice(i, i + 2000); sh.getRange(i + 2, 1, part.length, HEADERS.Ref.length).setValues(part); SpreadsheetApp.flush(); }
  clearRefCache_();
  console.log(`Imported ${data.length} SAP assets.`);
}

/* ===================== RESET (run from the editor before the real game) ===================== */
/**
 * Start the game fresh. Keeps: SAP data, approved accounts (points back to 0), the 2 starting locations.
 * Clears: finds, new items, points ledger, history, test locations, admin edits on Ref, test photos and exports.
 * A copy of the whole Sheet is saved first ("… backup …"); photos go to the Drive trash (restorable for 30 days).
 */
function resetGame(){
  ownerOnly_();
  withLock_(() => {
    const ss = ss_(), stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmm');
    const backup = DriveApp.getFileById(ss.getId()).makeCopy(ss.getName() + ' – backup ' + stamp);
    const rewrite = (name, keepFn, mapFn) => {
      const sh = ss.getSheetByName(name), last = sh.getLastRow(), w = HEADERS[name].length;
      if (last < 2) return 0;
      const rows = sh.getRange(2, 1, last - 1, w).getValues().map(r => r.map(cell_));
      const kept = rows.filter(keepFn || (() => false)).map(mapFn || (r => r));
      sh.getRange(2, 1, last - 1, w).clearContent();
      if (kept.length) sh.getRange(2, 1, kept.length, w).setNumberFormat('@').setValues(kept);
      return rows.length - kept.length;
    };
    const U = HEADERS.Users, ui = k => U.indexOf(k);
    const users = ss.getSheetByName('Users').getDataRange().getValues().slice(1).map(r => r.map(cell_));
    const keepNames = {}; users.forEach(r => { if (r[ui('Status')] === 'active') keepNames[r[ui('Username')]] = 1; });
    const removedUsers = rewrite('Users', r => r[ui('Status')] === 'active', r => { r[ui('Points')] = '0'; r[ui('Act Token')] = ''; r[ui('Reset Token')] = ''; r[ui('Reset Expires')] = ''; return r; });
    rewrite('Sessions', r => !!keepNames[r[1]]);
    const finds = rewrite('Updates'), news = rewrite('NewItems');
    rewrite('Points'); rewrite('History');
    const L = HEADERS.Locations;
    const removedLoc = rewrite('Locations', r => START_LOCATIONS.includes(r[0]), r => { r[L.indexOf('New')] = ''; r[L.indexOf('Active')] = 'Y'; return r; });
    const loc = tbl_('Locations');
    START_LOCATIONS.forEach(n => { if (!loc.rows.find(r => r.Location === n)) append_(loc, { Location:n, New:'', 'Added By':'setup', 'Added At':ts_(), Active:'Y' }); });
    const R = HEADERS.Ref, ri = k => R.indexOf(k);
    const removedRef = rewrite('Ref', r => r[ri('Added By')] === 'SAP import' || r[ri('Added By')] === '',
      r => { ['Status', 'Owner', 'Admin Note', 'Removed', 'Updated By', 'Updated At'].forEach(k => r[ri(k)] = ''); r[ri('Added By')] = 'SAP import'; return r; });
    let trashed = 0;
    ['PHOTO_FOLDER_ID', 'EXPORT_FOLDER_ID'].forEach(k => {
      const id = PropertiesService.getScriptProperties().getProperty(k); if (!id) return;
      try { const it = DriveApp.getFolderById(id).getFiles(); while (it.hasNext()){ it.next().setTrashed(true); trashed++; } } catch (e){}
    });
    clearRefCache_(); clearFoundCache_();
    console.log(`Reset done. Backup: ${backup.getUrl()}`);
    console.log(`Kept ${Object.keys(keepNames).length} approved accounts (points = 0), removed ${removedUsers} not-approved accounts.`);
    console.log(`Cleared ${finds} finds, ${news} new items, ${removedLoc} test locations, ${removedRef} admin-added assets; ${trashed} photo/export files moved to Drive trash.`);
  });
}
