/**
 * 分配法則ラボ（study-tools/distributive-law.html）専用のクラウド保存先。
 * スプレッドシートにバインドした GAS の Web アプリ。
 *
 * - 記録は「state」シートに JSON を分割して保存する（1セル 45,000 文字まで）。
 * - 前の版は「backup」シートに残す（1時間に1回まで・新しい 48 件まで）。
 * - 合言葉（SYNC_TOKEN）は、公開リポジトリに載せない別ファイル Secret.js で定義する。
 * - 合言葉を URL に載せないよう、読み書きはどちらも POST（body の t）で受ける。GET は ping だけ。
 * - 古い記録で新しい記録を上書きしない: 届いた記録が保存済みの「上位集合」でない時
 *   （xp が小さい・保存済みにある log の id／days の日付／cards の id が欠けている）は書かずに
 *   stale を返し、保存済みの記録を返す（アプリ側で合わせてから送り直す）。
 * - 返すのは学習の記録だけ。個人情報は扱わない。
 */
const CHUNK = 45000;
const BACKUP_KEEP = 48;
const BACKUP_MIN_INTERVAL_MS = 60 * 60 * 1000;

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function chunks_(raw) {
  const parts = [];
  for (let i = 0; i < raw.length; i += CHUNK) parts.push(raw.slice(i, i + CHUNK));
  return parts.length ? parts : [''];
}

function readRaw_() {
  const sh = sheet_('state');
  const last = sh.getLastColumn();
  if (last < 1) return '';
  return sh.getRange(1, 1, 1, last).getValues()[0].map(String).join('');
}

function writeRaw_(raw) {
  const sh = sheet_('state');
  const parts = chunks_(raw);
  const last = sh.getLastColumn();
  // 先に新しい値を書き、余った古い列だけ消す（途中で空になる時間を作らない）
  sh.getRange(1, 1, 1, parts.length).setNumberFormat('@').setValues([parts]);
  if (last > parts.length) sh.getRange(1, parts.length + 1, 1, last - parts.length).clearContent();
  sh.getRange(2, 1).setValue(new Date());
}

function backup_(raw, note) {
  if (!raw) return;
  const props = PropertiesService.getScriptProperties();
  const lastAt = Number(props.getProperty('lastBackupAt') || 0);
  if (Date.now() - lastAt < BACKUP_MIN_INTERVAL_MS) return;
  const sh = sheet_('backup');
  sh.insertRowBefore(1);
  const row = [new Date(), note].concat(chunks_(raw));
  sh.getRange(1, 1, 1, row.length).setNumberFormat('@').setValues([row.map(String)]);
  const n = sh.getLastRow();
  if (n > BACKUP_KEEP) sh.deleteRows(BACKUP_KEEP + 1, n - BACKUP_KEEP);
  props.setProperty('lastBackupAt', String(Date.now()));
}

function parse_(raw) {
  if (!raw) return null;
  try { const s = JSON.parse(raw); return (s && typeof s === 'object' && s.v === 1) ? s : null; } catch (e) { return null; }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function tokenOk_(t) {
  return typeof SYNC_TOKEN === 'string' && SYNC_TOKEN.length >= 24 && t === SYNC_TOKEN;
}

/* 届いた記録が保存済みの上位集合か（欠けている物が1つでもあれば false） */
function isSuperset_(inc, cur) {
  if (!cur) return true;
  if (typeof cur.xp === 'number' && inc.xp < cur.xp) return false;
  const ids = a => (Array.isArray(a) ? a : []).map(x => x && x.id).filter(Boolean);
  const incLog = {};
  ids(inc.log).forEach(id => { incLog[id] = 1; });
  if (ids(cur.log).some(id => !incLog[id])) return false;
  const keys = o => (o && typeof o === 'object') ? Object.keys(o) : [];
  const incDays = {}; keys(inc.days).forEach(k => { incDays[k] = 1; });
  if (keys(cur.days).some(k => !incDays[k])) return false;
  const incCards = {}; keys(inc.cards).forEach(k => { incCards[k] = 1; });
  if (keys(cur.cards).some(k => !incCards[k])) return false;
  return true;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'ping') {
    // 承認の確認だけ。記録は返さない。最初から使う API を全部呼んでおく（承認を1回で済ませるため）
    const o = { ok: true };
    try { sheet_('state'); PropertiesService.getScriptProperties().getProperty('lastBackupAt'); LockService.getScriptLock(); o.sheet = 'ok'; }
    catch (err) { o.sheet = 'ERR'; }
    try { o.authStatus = String(ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationStatus()); }
    catch (err) { o.authStatus = 'ERR'; }
    return out_(o);
  }
  return out_({ ok: false, error: 'action' });
}

function doPost(e) {
  let body = null;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'json' }); }
  if (!body || !tokenOk_(body.t)) return out_({ ok: false, error: 'token' });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return out_({ ok: false, error: 'busy' });
  try {
    const raw = readRaw_();
    const cur = parse_(raw);
    if (body.action === 'read') return out_({ ok: true, state: cur, size: raw.length });
    if (body.action !== 'write') return out_({ ok: false, error: 'action' });
    const inc = body.state;
    if (!inc || typeof inc !== 'object' || inc.v !== 1 || typeof inc.xp !== 'number') return out_({ ok: false, error: 'state' });
    if (!isSuperset_(inc, cur)) return out_({ ok: false, error: 'stale', state: cur });
    const next = JSON.stringify(inc);
    if (next === raw) return out_({ ok: true, same: true });
    backup_(raw, 'before write xp=' + (cur ? cur.xp : 'none'));
    writeRaw_(next);
    return out_({ ok: true, size: next.length });
  } finally {
    lock.releaseLock();
  }
}
