/* =========================================================
   保育園 物品管理台帳 — サーバー（Google Apps Script の Web アプリ）
   画面: hoikuen-bihin.html（GitHub Pages）。複数の端末で同じ台帳を使うための正本。

   公開範囲: 実行するユーザー＝自分（所有者）／アクセスできるユーザー＝全員（Google ログイン不要）
     使うのは別法人の保育園の職員で、この Google Workspace のアカウントを持たない。
     「Google ログイン必須」にしても園の区別はできず、「アクセスしているユーザーとして実行」は
     スプレッドシート自体の共有が要る（全園の行が見えてしまう）。そのため URL だけでは何もできず、
     すべての操作にこのアプリのログイン（メール＋パスワード → セッション）を求め、
     サーバーの中で「その人が所属する園」の表と写真フォルダだけを開く。

   保存場所: マイドライブ直下のフォルダ ROOT_NAME（介護の業務データとは別）
     管理用スプレッドシート（園・利用者・所属・セッション・招待コード）＋ 園ごとのスプレッドシートと写真フォルダ
   同期の規則（multi-device-sync）: 変えた項目だけを書く／在庫は増減／棚卸の確認と写真の並びは足し合わせ／
     削除は消した印（データは残す）／同じまとまり（batch）は一度だけ／番号（rev）はサーバーが振る
   ========================================================= */
const APP = 'hoikuen-bihin', PROTO = 1;
const ROOT_NAME = '保育園物品管理台帳_サーバーデータ（別法人用）';
const TABLES = ['items', 'assets', 'logs', 'stocktakes', 'settings', 'photos'];
// batch = the last write batch that touched the row (lets a retry after a half-finished write skip what is already there)
const COLS = ['id', 'data', 'deleted', 'rev', 'updated_at', 'updated_by', 'batch'];
const DEEP_FIELDS = ['checks', 'photoLabels', 'manuals'];   // manuals: one entry per model, registered from any device
// sha256 of the operator key (the key itself is kept outside the repository)
const OP_KEY_HASH = '7b712982a122e88ad57902e3d916ee1e96bd4b80d694db7bfb784a1fa760f35e';
const SESSION_DAYS = 30, PW_ITER = 500, MAX_CELL = 45000;
// 運営サポート: the operator joins one 園 as a visible member for a limited time (no backdoor into every 園)
const SUPPORT_DAYS = 90;
const MIME_OK = ['image/jpeg', 'image/png', 'image/webp', 'application/pdf'];

/* ---------------------------------------------------------------- entry points */
function doGet() {
  const out = { ok: true, app: APP, proto: PROTO };
  try {
    const info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    out.authStatus = String(info.getAuthorizationStatus());
  } catch (e) { out.authStatus = 'ERR'; }
  return json_(out);
}
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'bad_json', message: '送信内容を読めませんでした' }); }
  try { return json_(route_(req || {})); }
  catch (x) { return json_({ ok: false, error: x.code || 'server', message: String((x && x.message) || x).slice(0, 300) }); }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function fail_(code, msg) { const e = new Error(msg); e.code = code; throw e; }

function route_(r) {
  switch (r.action) {
    case 'version': return { ok: true, app: APP, proto: PROTO };
    case 'login': return login_(r);
    case 'redeem': return redeem_(r);
    case 'op': return op_(r);
  }
  const uid = session_(r.token);
  switch (r.action) {
    case 'me': return { ok: true, user: userInfo_(uid), orgs: myOrgs_(uid) };
    case 'logout': endSession_(r.token); return { ok: true };
    case 'apply': return apply_(uid, r.org, r.batch, r.ops, role_(uid, r.org));
    case 'pull': role_(uid, r.org); return pull_(r.org, r.since);
    case 'has': role_(uid, r.org); return has_(r.org, r.ids);
    case 'photoPut': role_(uid, r.org); return photoPut_(r);
    case 'photoGet': role_(uid, r.org); return photoGet_(r);
    case 'fetchPdf': role_(uid, r.org); return fetchPdf_(r);
    case 'members': admin_(uid, r.org); return { ok: true, members: members_(r.org) };
    case 'invite': return issueInvite_(uid, r.org, r.email, r.role || 'member', r.kind || 'join');
    case 'setMember': return setMember_(uid, r.org, r.user, r.role);
    case 'join': return join_(uid, r.code);
  }
  fail_('bad_action', '不明な操作です');
}

/* ---------------------------------------------------------------- storage helpers */
function props_() { return PropertiesService.getScriptProperties(); }
function root_() {
  const p = props_(); let id = p.getProperty('ROOT_ID');
  if (id) return DriveApp.getFolderById(id);
  const f = DriveApp.createFolder(ROOT_NAME);
  p.setProperty('ROOT_ID', f.getId()); return f;
}
function reg_() {
  const p = props_(); let id = p.getProperty('REG_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('保育園物品管理台帳_管理（園・利用者）');
  DriveApp.getFileById(ss.getId()).moveTo(root_());
  p.setProperty('REG_ID', ss.getId()); return ss;
}
const REG_SHEETS = {
  orgs: ['org_id', 'name', 'sheet_id', 'folder_id', 'created_at'],
  users: ['uid', 'email', 'pw', 'created_at', 'kind'],
  members: ['org_id', 'uid', 'role', 'created_at', 'expires'],
  sessions: ['token_hash', 'uid', 'expires'],
  invites: ['id', 'org_id', 'email', 'role', 'kind', 'code_hash', 'created_by', 'created_at', 'expires', 'used_at', 'used_by', 'attempts']
};
function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@'); sh.getRange(1, 1, 1, header.length).setValues([header]); sh.setFrozenRows(1); }
  return sh;
}
// a new sheet has 1000 rows and getRange() past the end throws: grow it first (new rows also plain text)
function ensureRows_(sh, n) {
  const max = sh.getMaxRows();
  if (max >= n) return;
  const add = n - max + 500;
  sh.insertRowsAfter(max, add);
  sh.getRange(max + 1, 1, add, sh.getMaxColumns()).setNumberFormat('@');
}
function appendSafe_(sh, row) { const r = sh.getLastRow() + 1; ensureRows_(sh, r); sh.getRange(r, 1, 1, row.length).setValues([row]); }
// whole sheet as objects (row = 1-based sheet row)
function rows_(sh) {
  const v = sh.getDataRange().getValues(); const h = v[0] || [];
  const out = [];
  for (let i = 1; i < v.length; i++) { const o = { _row: i + 1 }; h.forEach((k, j) => { o[k] = v[i][j]; }); out.push(o); }
  return out;
}
function regRows_(name) { return rows_(sheet_(reg_(), name, REG_SHEETS[name])); }
// orgs / members are read on every request: keep a 60-second copy, dropped on every change
// the copy carries a generation number; any change bumps it, so a copy read just before a change is never used after it
function regCached_(name) {
  const c = CacheService.getScriptCache(), k = 'reg_' + name, gen = props_().getProperty('gen_' + name) || '0', hit = c.get(k);
  if (hit) { try { const o = JSON.parse(hit); if (o.g === gen) return o.v; } catch (e) {} }
  const v = regRows_(name); try { c.put(k, JSON.stringify({ g: gen, v: v }), 60); } catch (e) {}
  return v;
}
function regChanged_(name) { props_().setProperty('gen_' + name, String(Date.now()) + Math.random().toString(36).slice(2, 6)); CacheService.getScriptCache().remove('reg_' + name); }
function regAppend_(name, obj) { const h = REG_SHEETS[name]; appendSafe_(sheet_(reg_(), name, h), h.map(k => obj[k] == null ? '' : obj[k])); regChanged_(name); }
function regUpdate_(name, row, obj) { const h = REG_SHEETS[name]; sheet_(reg_(), name, h).getRange(row, 1, 1, h.length).setValues([h.map(k => obj[k] == null ? '' : obj[k])]); regChanged_(name); }
function regDelete_(name, row) { sheet_(reg_(), name, REG_SHEETS[name]).deleteRow(row); regChanged_(name); }
function org_(orgId) {
  const o = regCached_('orgs').find(x => x.org_id === orgId);
  if (!o) fail_('forbidden', 'この園のメンバーではありません');
  return o;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) fail_('busy', '混み合っています。少し待ってからもう一度お試しください');
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------------------------------------------------------- crypto */
function hex_(bytes) { return bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }
function sha256_(s) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)); }
// stored hashes / ids start with a letter so that Sheets never reads them as numbers (e.g. 123e45…)
const hid_ = s => 'h' + sha256_(s);
function randomToken_(n) {   // n*16 hex chars from Utilities.getUuid (crypto-strong v4 UUIDs)
  let s = ''; for (let i = 0; i < n; i++) s += Utilities.getUuid().replace(/-/g, ''); return s;
}
function hashPw_(pw, salt, iter) { let h = sha256_(salt + ':' + pw); for (let i = 1; i < iter; i++) h = sha256_(h + salt); return h; }
function makePw_(pw) { const salt = randomToken_(1); return ['v1', PW_ITER, salt, hashPw_(pw, salt, PW_ITER)].join('$'); }
function checkPw_(pw, stored) {
  const p = String(stored || '').split('$'); if (p[0] !== 'v1') return false;
  const h = hashPw_(pw, p[2], Number(p[1]));
  let d = 0; for (let i = 0; i < h.length; i++) d |= h.charCodeAt(i) ^ (p[3] || '').charCodeAt(i);
  return d === 0 && h.length === (p[3] || '').length;
}
const normEmail_ = e => String(e || '').trim().toLowerCase();
const okEmail_ = e => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e) && e.length <= 254;
const normCode_ = c => String(c || '').replace(/[^0-9A-Za-z]/g, '').toUpperCase();

/* ---------------------------------------------------------------- sessions */
function login_(r) {
  const email = normEmail_(r.email), pw = String(r.password || '');
  const cache = CacheService.getScriptCache(), fk = 'fail_' + sha256_(email);
  withLock_(() => {   // count the attempt first: many guesses sent at once still cannot exceed 10 per 15 minutes
    const n = Number(cache.get(fk) || 0);
    if (n >= 10) fail_('locked', 'ログインの失敗が続いたため、15分ほど待ってからお試しください（ログイン中のほかの端末はそのまま使えます）');
    cache.put(fk, String(n + 1), 900);
  });
  const u = regRows_('users').find(x => x.email === email);
  if (!u || !checkPw_(pw, u.pw)) fail_('login', 'メールアドレスかパスワードが違います');
  cache.remove(fk);
  const token = randomToken_(2);
  withLock_(() => regAppend_('sessions', { token_hash: hid_(token), uid: u.uid, expires: new Date(Date.now() + SESSION_DAYS * 864e5).toISOString() }));
  return { ok: true, token: token, user: { uid: u.uid, email: u.email }, orgs: myOrgs_(u.uid) };
}
function session_(token) {
  if (!token || typeof token !== 'string' || token.length < 32) fail_('auth', 'ログインしてください');
  const th = hid_(token), cache = CacheService.getScriptCache();
  const c = cache.get('s_' + th);
  if (c) { const i = c.indexOf('|'), uid = c.slice(0, i); if ((props_().getProperty('sgen_' + uid) || '0') === c.slice(i + 1)) return uid; }
  const s = regRows_('sessions').find(x => x.token_hash === th);
  if (!s || new Date(s.expires) < new Date()) fail_('auth', 'ログインの期限が切れました。もう一度ログインしてください');
  cache.put('s_' + th, s.uid + '|' + (props_().getProperty('sgen_' + s.uid) || '0'), 600);
  return s.uid;
}
function endSession_(token) {
  const th = hid_(String(token || ''));
  CacheService.getScriptCache().remove('s_' + th);
  withLock_(() => { const s = regRows_('sessions').find(x => x.token_hash === th); if (s) regUpdate_('sessions', s._row, { token_hash: 'ended', uid: s.uid, expires: new Date(0).toISOString() }); });
}
function userInfo_(uid) { const u = regRows_('users').find(x => x.uid === uid); return u ? { uid: u.uid, email: u.email } : null; }
// a support membership counts only until its expiry (the server refuses afterwards, whatever the device has cached)
const live_ = m => m.role !== 'support' || (!!m.expires && new Date(m.expires) > new Date());
function myOrgs_(uid) {
  const orgs = regCached_('orgs');
  return regCached_('members').filter(m => m.uid === uid && live_(m)).map(m => { const o = orgs.find(x => x.org_id === m.org_id); return o ? { org_id: o.org_id, name: o.name, role: m.role } : null; }).filter(Boolean);
}
function role_(uid, orgId) {
  const m = regCached_('members').find(x => x.uid === uid && x.org_id === orgId);
  if (!m) fail_('forbidden', 'この園のメンバーではありません');
  if (!live_(m)) fail_('forbidden', '運営サポートの期限が過ぎました。運営者に延長を頼んでください');
  return m.role;
}
function admin_(uid, orgId) { if (role_(uid, orgId) !== 'admin') fail_('forbidden', '管理者だけができる操作です'); }

/* ---------------------------------------------------------------- invites / accounts */
function newInvite_(orgId, email, role, kind, by) {
  const A = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, randomToken_(2));
  let c = ''; for (let i = 0; i < 12; i++) c += A[(bytes[i] & 255) % 32];
  const now = new Date();
  regRows_('invites').forEach(x => { if (x.org_id === orgId && x.email === email && x.kind === kind && !x.used_at && new Date(x.expires) > now) { x.expires = now.toISOString(); regUpdate_('invites', x._row, x); } });
  const expires = new Date(now.getTime() + (kind === 'reset' ? 1 : 7) * 864e5).toISOString();
  regAppend_('invites', { id: 'i' + randomToken_(1), org_id: orgId, email: email, role: role, kind: kind, code_hash: hid_(c), created_by: by || '', created_at: now.toISOString(), expires: expires, used_at: '', used_by: '', attempts: 0 });
  return { code: c.slice(0, 4) + '-' + c.slice(4, 8) + '-' + c.slice(8), email: email, kind: kind, expires_at: expires };
}
function issueInvite_(uid, orgId, email, role, kind) {
  admin_(uid, orgId);
  email = normEmail_(email);
  if (!okEmail_(email)) fail_('bad_email', 'メールアドレスの形が正しくありません');
  if (['admin', 'member'].indexOf(role) < 0 || ['join', 'reset'].indexOf(kind) < 0) fail_('bad_input', 'bad role/kind');
  return withLock_(() => {
    if (kind === 'reset') {
      const u = regRows_('users').find(x => x.email === email);
      const mem = regRows_('members');
      if (!u || !mem.some(m => m.uid === u.uid && m.org_id === orgId)) fail_('bad_input', 'この園のメンバーではないため、パスワードの再設定コードは出せません');
      if (mem.some(m => m.uid === u.uid && m.role === 'support')) fail_('forbidden', '運営サポートのパスワードは、園からは再設定できません');
      // an admin may only reset accounts that belong exclusively to 園 they administer
      if (mem.some(m => m.uid === u.uid && !mem.some(a => a.org_id === m.org_id && a.uid === uid && a.role === 'admin'))) fail_('forbidden', 'ほかの園にも所属している人のため、ここでは再設定できません');
    }
    return Object.assign({ ok: true }, newInvite_(orgId, email, role, kind, uid));
  });
}
// wrong codes only slow down that email address for 15 minutes; they never burn the valid codes
// (otherwise anyone who knows the address could make the real invitation unusable)
function findInvite_(email, code, kind) {
  const cache = CacheService.getScriptCache(), fk = 'ifail_' + sha256_(email), fails = Number(cache.get(fk) || 0);
  if (fails >= 10) fail_('locked', 'コードの間違いが続いたため、15分ほど待ってからお試しください');
  const now = new Date(), h = hid_(normCode_(code));
  const inv = regRows_('invites').find(x => x.email === email && !x.used_at && new Date(x.expires) > now && x.code_hash === h && (!kind || x.kind === kind));
  if (!inv) cache.put(fk, String(fails + 1), 900); else cache.remove(fk);
  return inv || null;
}
// reset codes are checked again when used: the person must still belong to that 園 and only to 園 the issuer administers
function resetStillAllowed_(inv, u) {
  if (!inv.created_by) return true;   // issued by the operator
  const mem = regRows_('members');
  if (!mem.some(m => m.uid === u.uid && m.org_id === inv.org_id)) return false;
  if (mem.some(m => m.uid === u.uid && m.role === 'support')) return false;
  return !mem.some(m => m.uid === u.uid && !mem.some(a => a.org_id === m.org_id && a.uid === inv.created_by && a.role === 'admin'));
}
function redeem_(r) {
  const email = normEmail_(r.email), code = normCode_(r.code), pw = String(r.password || '');
  if (!okEmail_(email)) fail_('email', 'メールアドレスの形が正しくありません');
  if (code.length !== 12) fail_('code', 'コードは12文字です（ハイフンはあってもなくても構いません）');
  if (pw.length < 8 || pw.length > 200) fail_('password', 'パスワードは8文字以上にしてください');
  return withLock_(() => {
    const inv = findInvite_(email, code);
    if (!inv) fail_('invalid', 'コードが正しくないか、期限が切れています。園の管理者に新しいコードを出してもらってください');
    const users = regRows_('users'); let u = users.find(x => x.email === email);
    if (inv.kind === 'join') {
      if (u) fail_('exists', 'このメールアドレスはすでに登録されています。ログインしてから「設定 → 招待コードで別の園に参加」でコードを入れてください');
      if (inv.role === 'support') usersHeader_();
      u = { uid: 'u' + randomToken_(1), email: email, pw: makePw_(pw), created_at: new Date().toISOString(), kind: inv.role === 'support' ? 'support' : '' };
      regAppend_('users', u);
      if (!regRows_('members').some(m => m.uid === u.uid && m.org_id === inv.org_id)) addMember_(inv.org_id, u.uid, inv.role);
    } else {
      if (!u || !resetStillAllowed_(inv, u)) fail_('invalid', 'コードが正しくないか、期限が切れています');
      u.pw = makePw_(pw); regUpdate_('users', u._row, u);
      // a new password ends every existing session of that person
      props_().setProperty('sgen_' + u.uid, String(Date.now()));   // cached copies of the old sessions stop working at once
      regRows_('sessions').forEach(s => { if (s.uid === u.uid && s.token_hash !== 'ended') { CacheService.getScriptCache().remove('s_' + s.token_hash); regUpdate_('sessions', s._row, { token_hash: 'ended', uid: s.uid, expires: new Date(0).toISOString() }); } });
    }
    inv.used_at = new Date().toISOString(); inv.used_by = u.uid; regUpdate_('invites', inv._row, inv);
    return { ok: true, kind: inv.kind };
  });
}
function join_(uid, code) {
  return withLock_(() => {
    const u = regRows_('users').find(x => x.uid === uid);
    const inv = u && findInvite_(u.email, code, 'join');
    if (!inv) fail_('invalid', 'コードが正しくないか、期限が切れています');
    if (!regRows_('members').some(m => m.uid === uid && m.org_id === inv.org_id)) addMember_(inv.org_id, uid, inv.role);
    inv.used_at = new Date().toISOString(); inv.used_by = uid; regUpdate_('invites', inv._row, inv);
    return { ok: true, org_id: inv.org_id };
  });
}
function supportExpiry_(days) { return new Date(Date.now() + (days || SUPPORT_DAYS) * 864e5).toISOString(); }
function usersHeader_() { const sh = sheet_(reg_(), 'users', REG_SHEETS.users); if (sh.getRange(1, 5).getValue() !== 'kind') sh.getRange(1, 5).setValue('kind'); }
function membersHeader_() { const sh = sheet_(reg_(), 'members', REG_SHEETS.members); if (sh.getRange(1, 5).getValue() !== 'expires') sh.getRange(1, 5).setValue('expires'); }
function addMember_(orgId, uid, role) {
  membersHeader_();
  regAppend_('members', { org_id: orgId, uid: uid, role: role, created_at: new Date().toISOString(), expires: role === 'support' ? supportExpiry_() : '' });
}
function members_(orgId) {
  const users = regRows_('users');
  return regRows_('members').filter(m => m.org_id === orgId).map(m => { const u = users.find(x => x.uid === m.uid); return { user_id: m.uid, email: u ? u.email : '', role: m.role, created_at: m.created_at, expires: m.role === 'support' ? m.expires : '' }; });
}
function setMember_(uid, orgId, target, role) {
  admin_(uid, orgId);
  if (role != null && ['admin', 'member'].indexOf(role) < 0) fail_('bad_input', 'bad role');
  return withLock_(() => {
    const mem = regRows_('members'), m = mem.find(x => x.org_id === orgId && x.uid === target);
    if (!m) fail_('bad_input', 'この園のメンバーではありません');
    if (m.role === 'support' && role != null) fail_('bad_input', '運営サポートの役割は変えられません（外すことはできます）');
    if ((role == null || role === 'member') && m.role === 'admin' && mem.filter(x => x.org_id === orgId && x.role === 'admin').length <= 1) fail_('bad_input', '管理者が1人もいなくなるため変更できません');
    if (role == null) {
      regDelete_('members', m._row);
      const u = regRows_('users').find(x => x.uid === target), now = new Date();
      if (u) regRows_('invites').forEach(x => { if (x.org_id === orgId && x.email === u.email && !x.used_at && new Date(x.expires) > now) { x.expires = now.toISOString(); regUpdate_('invites', x._row, x); } });
    }
    else { m.role = role; regUpdate_('members', m._row, m); }
    return { ok: true };
  });
}

/* ---------------------------------------------------------------- operator (key kept outside the repository) */
function op_(r) {
  if (!r.key || sha256_(String(r.key)) !== OP_KEY_HASH) fail_('forbidden', 'forbidden');
  if (r.op === 'createOrg') return withLock_(() => {
    const name = String(r.name || '').trim(), email = normEmail_(r.adminEmail);
    if (!name || name.length > 100 || !okEmail_(email)) fail_('bad_input', 'bad name/email');
    const orgId = Utilities.getUuid();
    const folder = root_().createFolder('園_' + name + '_' + orgId.slice(0, 8));
    const ss = SpreadsheetApp.create('台帳_' + name);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    TABLES.forEach(t => sheet_(ss, t, COLS)); sheet_(ss, 'applied', ['batch', 'at']); sheet_(ss, 'files', ['key', 'file_id']);
    const first = ss.getSheetByName('Sheet1') || ss.getSheetByName('シート1'); if (first) ss.deleteSheet(first);
    const photos = folder.createFolder('写真');
    regAppend_('orgs', { org_id: orgId, name: name, sheet_id: ss.getId(), folder_id: photos.getId(), created_at: new Date().toISOString() });
    return Object.assign({ ok: true, org_id: orgId }, newInvite_(orgId, email, 'admin', 'join', ''));
  });
  // recovery by the operator (e.g. an address registered by someone other than its owner)
  if (r.op === 'issueReset') return withLock_(() => {
    const email = normEmail_(r.email), u = regRows_('users').find(x => x.email === email);
    if (!u) fail_('bad_input', 'no such user');
    const m = regRows_('members').find(x => x.uid === u.uid);
    return Object.assign({ ok: true }, newInvite_(m ? m.org_id : '', email, 'member', 'reset', ''));
  });
  // 運営サポート: add the operator account to one 園 (existing account → member now; new → a join code), or set its expiry
  if (r.op === 'addSupport') return withLock_(() => {
    const email = normEmail_(r.email), o = regRows_('orgs').find(x => x.org_id === r.org);
    if (!o || !okEmail_(email)) fail_('bad_input', 'bad org/email');
    const days = Math.min(365, Math.max(1, Number(r.days) || SUPPORT_DAYS));
    const u = regRows_('users').find(x => x.email === email);
    if (u) {
      // only an account created from a support code (never one somebody else may have registered with this address)
      if (u.kind !== 'support') fail_('bad_input', 'this address belongs to a regular account; a support account must be created from a support code');
      membersHeader_();
      const m = regRows_('members').find(x => x.org_id === o.org_id && x.uid === u.uid);
      if (m && m.role !== 'support') fail_('bad_input', 'already a regular member');
      const row = { org_id: o.org_id, uid: u.uid, role: 'support', created_at: m ? m.created_at : new Date().toISOString(), expires: supportExpiry_(days) };
      if (m) regUpdate_('members', m._row, row); else regAppend_('members', row);
      return { ok: true, member: true, expires: row.expires };
    }
    if (r.days && days !== SUPPORT_DAYS) fail_('bad_input', 'no account yet: the period starts at ' + SUPPORT_DAYS + ' days when the code is used; change it afterwards with setSupportExpiry');
    return Object.assign({ ok: true, member: false }, newInvite_(o.org_id, email, 'support', 'join', ''));
  });
  if (r.op === 'setSupportExpiry') return withLock_(() => {
    const u = regRows_('users').find(x => x.email === normEmail_(r.email));
    const m = u && regRows_('members').find(x => x.org_id === r.org && x.uid === u.uid && x.role === 'support');
    if (!m || isNaN(new Date(r.expires))) fail_('bad_input', 'no such support member / bad date');
    m.expires = new Date(r.expires).toISOString(); regUpdate_('members', m._row, m);
    return { ok: true, expires: m.expires };
  });
  if (r.op === 'listOrgs') return { ok: true, orgs: regRows_('orgs').map(o => ({ org_id: o.org_id, name: o.name, members: regRows_('members').filter(m => m.org_id === o.org_id).length })) };
  // test cleanup only: 園 whose name starts with テスト園 (moved to the trash, restorable for 30 days)
  if (r.op === 'dropTestOrg') return withLock_(() => {
    const o = regRows_('orgs').find(x => x.org_id === r.org);
    if (!o || String(o.name).indexOf('テスト園') !== 0) fail_('forbidden', 'only test 園 can be dropped');
    DriveApp.getFileById(o.sheet_id).getParents().next().setTrashed(true);
    regRows_('members').filter(m => m.org_id === o.org_id).reverse().forEach(m => regDelete_('members', m._row));
    regDelete_('orgs', o._row);
    return { ok: true };
  });
  fail_('bad_action', 'bad op');
}

/* ---------------------------------------------------------------- merge (same rules as the client's mergeJS) */
const isObj_ = v => v != null && typeof v === 'object' && !Array.isArray(v);
function merge_(cur, set, inc, arr) {
  const nd = Object.assign({}, cur || {});
  if (isObj_(set)) Object.keys(set).forEach(k => {
    const v = set[k];
    if (!(DEEP_FIELDS.indexOf(k) >= 0 && isObj_(v))) { nd[k] = v; return; }
    const o = Object.assign({}, isObj_(nd[k]) ? nd[k] : {});
    // manuals: each model's entry merges field by field (a URL from one device and a PDF from another both stay)
    Object.keys(v).forEach(kk => { o[kk] = (k === 'manuals' && isObj_(v[kk]) && isObj_(o[kk])) ? Object.assign({}, o[kk], v[kk]) : v[kk]; });
    nd[k] = o;
  });
  if (isObj_(inc)) Object.keys(inc).forEach(k => {
    const v = inc[k]; if (typeof v !== 'number' || !isFinite(v)) return;
    const b = typeof nd[k] === 'number' ? nd[k] : (/^-?\d+(\.\d+)?$/.test(String(nd[k] == null ? '' : nd[k])) ? Number(nd[k]) : 0);
    nd[k] = b + v;
  });
  if (isObj_(arr)) Object.keys(arr).forEach(k => {
    const v = arr[k]; if (!isObj_(v)) return;
    const rem = Array.isArray(v.rem) ? v.rem.map(x => JSON.stringify(x)) : [];
    const res = (Array.isArray(nd[k]) ? nd[k] : []).filter(e => rem.indexOf(JSON.stringify(e)) < 0);
    (Array.isArray(v.add) ? v.add : []).forEach(e => { if (!res.some(x => JSON.stringify(x) === JSON.stringify(e))) res.push(e); });
    if (Array.isArray(v.order)) {   // the editor's order first, then anything only the server had (added elsewhere)
      const ord = v.order.map(x => JSON.stringify(x)), pos = e => { const i = ord.indexOf(JSON.stringify(e)); return i < 0 ? ord.length : i; };
      const idx = res.map((e, i) => [e, pos(e), i]);
      idx.sort((a, b) => a[1] - b[1] || a[2] - b[2]); nd[k] = idx.map(x => x[0]);
    } else nd[k] = res;
  });
  delete nd.id;
  return nd;
}

/* ---------------------------------------------------------------- data API */
function orgSS_(orgId) { return SpreadsheetApp.openById(org_(orgId).sheet_id); }
function revKey_(orgId) { return 'rev_' + orgId; }
function loadTable_(ss, t) {
  const sh = sheet_(ss, t, COLS), v = sh.getDataRange().getValues(), map = {};
  const list = [];
  for (let i = 1; i < v.length; i++) {
    let data; try { data = JSON.parse(v[i][1]); } catch (e) { data = {}; }
    const rec = { id: String(v[i][0]), data: isObj_(data) ? data : {}, deleted: String(v[i][2]).toUpperCase() === 'TRUE', rev: Number(v[i][3]) || 0, updated_at: v[i][4], updated_by: v[i][5], batch: String(v[i][6] || ''), _i: list.length };
    rec.b0 = rec.batch;
    list.push(rec); map[rec.id] = rec;
  }
  return { sh: sh, list: list, map: map, dirty: false };
}
function apply_(uid, orgId, batch, ops, role) {
  if (typeof batch !== 'string' || batch.length < 1 || batch.length > 100) fail_('bad_input', 'bad batch id');
  if (!Array.isArray(ops) || ops.length > 1000) fail_('bad_input', 'bad ops');
  ops.forEach(op => {   // validate everything before writing anything (a bad batch changes nothing)
    if (!op || TABLES.indexOf(op.t) < 0) fail_('bad_input', 'bad table ' + (op && op.t));
    if (typeof op.id !== 'string' || op.id.length < 1 || op.id.length > 64) fail_('bad_input', 'bad id');
    if (op.k != null && ['patch', 'delete', 'create'].indexOf(op.k) < 0) fail_('bad_input', 'bad kind');
    if ((op.set != null && !isObj_(op.set)) || (op.inc != null && !isObj_(op.inc)) || (op.arr != null && !isObj_(op.arr))) fail_('bad_input', 'bad patch');
  });
  return withLock_(() => {
    const ss = orgSS_(orgId);
    const ap = sheet_(ss, 'applied', ['batch', 'at']);
    const done = ap.getLastRow() > 1 ? ap.getRange(2, 1, ap.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
    if (done.indexOf(batch) >= 0) return { ok: true, dup: true };
    const T = {}; const tab = t => T[t] || (T[t] = loadTable_(ss, t));
    const p = props_(); let rev = Number(p.getProperty(revKey_(orgId)) || 0);
    const now = new Date().toISOString();
    const touch = (t, rec) => { rev++; rec.rev = rev; rec.updated_at = now; rec.updated_by = uid; rec.batch = batch; tab(t).dirty = true; };
    ops.forEach(op => {
      const tb = tab(op.t), cur = tb.map[op.id];
      // already written by an earlier, half-finished run of this same batch: do not apply it twice (stock deltas)
      if (cur && cur.b0 === batch) return;
      if (op.k === 'delete') {
        if (!cur || cur.deleted) return;
        // an item that another device has just given a live unit stays (its name would otherwise vanish)
        if (op.t === 'items' && tab('assets').list.some(a => !a.deleted && a.data.itemId === op.id)) return;
        cur.deleted = true; touch(op.t, cur); return;
      }
      // 'create' (migration / backup): only when the id never existed — a record deleted on purpose stays deleted
      if (op.k === 'create' && cur) return;
      const nd = merge_(cur ? cur.data : {}, op.set, op.inc, op.arr);
      if (op.t === 'logs' && role === 'support') nd.support = true;   // written during 運営サポート
      if (JSON.stringify(nd).length > MAX_CELL) fail_('too_large', '1件のデータが大きすぎます');
      if (cur) { cur.data = nd; cur.deleted = false; touch(op.t, cur); }   // a patch revives a tombstone
      else { const rec = { id: op.id, data: nd, deleted: false, _i: tb.list.length }; tb.list.push(rec); tb.map[op.id] = rec; touch(op.t, rec); }
      const revive = (t, id) => { const r2 = id && tab(t).map[id]; if (r2 && r2.deleted) { r2.deleted = false; touch(t, r2); } };
      if (op.t === 'assets') revive('items', nd.itemId);
      if (op.t === 'logs') { revive('assets', nd.assetId); revive('items', nd.itemId); }
    });
    Object.keys(T).forEach(t => {
      const tb = T[t]; if (!tb.dirty || !tb.list.length) return;
      ensureRows_(tb.sh, tb.list.length + 1);
      tb.sh.getRange(2, 1, tb.list.length, COLS.length).setValues(tb.list.map(r => [r.id, JSON.stringify(r.data), r.deleted, r.rev, r.updated_at || '', r.updated_by || '', r.batch || '']));
    });
    appendSafe_(ap, [batch, now]);
    SpreadsheetApp.flush();
    p.setProperty(revKey_(orgId), String(rev));   // rows above this number are only shown once everything is written
    return { ok: true, n: ops.length, rev: rev };
  });
}
function pull_(orgId, since) {
  since = Number(since) || 0;
  const committed = Number(props_().getProperty(revKey_(orgId)) || 0);
  org_(orgId);
  if (since >= committed) return { ok: true, rev: committed, rows: {} };
  const ss = orgSS_(orgId), rows = {};
  TABLES.forEach(t => {
    rows[t] = loadTable_(ss, t).list.filter(r => r.rev > since && r.rev <= committed).map(r => ({ id: r.id, data: r.data, deleted: r.deleted, rev: r.rev }));
  });
  return { ok: true, rev: committed, rows: rows };
}

// which ids exist on the server (deleted or not) — used to confirm a migration really arrived
function has_(orgId, ids) {
  if (!isObj_(ids)) fail_('bad_input', 'bad ids');
  const ss = orgSS_(orgId), out = {};
  Object.keys(ids).forEach(t => {
    if (TABLES.indexOf(t) < 0 || !Array.isArray(ids[t])) fail_('bad_input', 'bad table');
    const m = loadTable_(ss, t).map;
    out[t] = ids[t].filter(id => typeof id === 'string' && m[id]);
  });
  return { ok: true, has: out };
}

/* ---------------------------------------------------------------- photos / PDF (private Drive folder per 園) */
function fileKey_(r) {
  if (typeof r.id !== 'string' || !/^[A-Za-z0-9_-]{1,64}$/.test(r.id)) fail_('bad_input', 'bad photo id');
  if (r.which !== 'main' && r.which !== 'thumb') fail_('bad_input', 'bad which');
  return r.id + (r.which === 'thumb' ? '_t' : '');
}
function photoPut_(r) {
  const key = fileKey_(r);
  if (MIME_OK.indexOf(r.mime) < 0) fail_('bad_input', 'この種類のファイルは保存できません');
  const bytes = Utilities.base64Decode(String(r.b64 || ''));
  if (!bytes.length || bytes.length > 15 * 1048576) fail_('too_large', 'ファイルが大きすぎます（15MBまで）');
  const o = org_(r.org);
  return withLock_(() => {
    const ss = SpreadsheetApp.openById(o.sheet_id), sh = sheet_(ss, 'files', ['key', 'file_id']);
    const ex = rows_(sh).find(x => x.key === key);
    if (ex) return { ok: true, same: true };   // the same photo id is never re-used with other content
    const f = DriveApp.getFolderById(o.folder_id).createFile(Utilities.newBlob(bytes, r.mime, key));
    appendSafe_(sh, [key, f.getId()]);
    return { ok: true };
  });
}
function photoGet_(r) {
  const key = fileKey_(r), o = org_(r.org);
  const ex = rows_(sheet_(SpreadsheetApp.openById(o.sheet_id), 'files', ['key', 'file_id'])).find(x => x.key === key);
  if (!ex) fail_('not_found', '写真がまだサーバーにありません');
  const b = DriveApp.getFileById(ex.file_id).getBlob();
  return { ok: true, mime: b.getContentType(), b64: Utilities.base64Encode(b.getBytes()) };
}

// keep a copy of a manual PDF that is on the web (in case the link breaks later). Stored like an attached PDF.
function fetchPdf_(r) {
  const key = fileKey_({ id: r.id, which: 'main' });
  const url = String(r.url || '');
  if (!/^https?:\/\/[^\s]{4,2000}$/.test(url)) fail_('bad_input', 'URL の形が正しくありません');
  let res;
  try { res = UrlFetchApp.fetch(url, { followRedirects: true, muteHttpExceptions: true, validateHttpsCertificates: true }); }
  catch (e) { fail_('fetch', 'そのURLを開けませんでした'); }
  if (res.getResponseCode() !== 200) fail_('fetch', 'そのURLを開けませんでした（' + res.getResponseCode() + '）');
  const bytes = res.getContent();
  const isPdf = bytes.length > 4 && bytes[0] === 37 && bytes[1] === 80 && bytes[2] === 68 && bytes[3] === 70;   // %PDF
  if (!isPdf) fail_('not_pdf', 'PDF ではないため保存できません（説明書のページのURLは、そのまま登録してください）');
  if (bytes.length > 15 * 1048576) fail_('too_large', 'PDF が大きすぎます（15MBまで）');
  const o = org_(r.org);
  return withLock_(() => {
    const ss = SpreadsheetApp.openById(o.sheet_id), sh = sheet_(ss, 'files', ['key', 'file_id']);
    if (!rows_(sh).some(x => x.key === key)) {
      const f = DriveApp.getFolderById(o.folder_id).createFile(Utilities.newBlob(bytes, 'application/pdf', key));
      appendSafe_(sh, [key, f.getId()]);
    }
    return { ok: true, size: bytes.length };
  });
}

/* used once from the editor/authorize flow: touches every service the app needs (gas-deploy-auth 鉄則2) */
function setup() { root_(); reg_(); UrlFetchApp.getRequest('https://www.google.com/'); return { ok: true }; }
