/**
 * 會議記錄系統（組別版）：Google 試算表資料庫
 * 貼到這份試算表的「擴充功能 → Apps Script」，取代原本全部的程式，
 * 再到「部署 → 管理部署作業」→ 鉛筆 →「版本：新版本」→ 部署（網址不會變）。
 */
const DB_SHEET = '_db';
const LEGACY_SHEET = '_data';
const KINDS = ['group', 'meeting', 'report'];

function doGet(e) {
  try {
    const all = readAll_();
    return json_({ ok: true, groups: all.group, meetings: all.meeting, reports: all.report });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'save') saveOne_(req.kind, req.item);
    else if (req.action === 'saveMany') (req.items || []).forEach(it => saveOne_(req.kind, it));
    else if (req.action === 'delete') del_(req.kind, req.id);
    else return json_({ ok: false, error: 'unknown action' });
    rebuildViews_();
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function db_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DB_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DB_SHEET);
    sh.getRange('A:B').setNumberFormat('@');
    sh.getRange(1, 1, 1, 4).setValues([['kind', 'id', 'updatedAt', 'json']]);
    migrateLegacy_(sh);
    sh.hideSheet();
  }
  return sh;
}

/* 舊版（沒有組別）的會議資料搬過來 */
function migrateLegacy_(sh) {
  const old = SpreadsheetApp.getActive().getSheetByName(LEGACY_SHEET);
  if (!old) return;
  const n = old.getLastRow();
  if (n < 2) return;
  const rows = old.getRange(2, 1, n - 1, 3).getValues()
    .filter(r => r[0])
    .map(r => ['meeting', String(r[0]), r[1], r[2]]);
  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
}

function rows_() {
  const sh = db_();
  const n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, 4).getValues();
}

function readAll_() {
  const out = { group: [], meeting: [], report: [] };
  rows_().forEach(r => {
    if (!out[r[0]]) return;
    try { out[r[0]].push(JSON.parse(r[3])); } catch (e) {}
  });
  return out;
}

function saveOne_(kind, item) {
  if (KINDS.indexOf(kind) < 0) throw new Error('bad kind');
  if (!item || !item.id) throw new Error('missing id');
  const text = JSON.stringify(item);
  if (text.length > 49000) throw new Error('內容太長，請把這場會議分成兩場記錄');
  const sh = db_();
  const rows = rows_();
  const i = rows.findIndex(r => r[0] === kind && String(r[1]) === String(item.id));
  const row = [kind, String(item.id), Number(item.updatedAt) || Date.now(), text];
  if (i >= 0) {
    if (row[2] < (Number(rows[i][2]) || 0)) return; // 比較舊的版本，不覆蓋
    sh.getRange(i + 2, 1, 1, 4).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}

function del_(kind, id) {
  const sh = db_();
  const rows = rows_();
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    let hit = r[0] === kind && String(r[1]) === String(id);
    if (!hit && kind === 'meeting' && r[0] === 'report') {
      try { hit = String(JSON.parse(r[3]).itemKey || '').indexOf(String(id) + ':') === 0; } catch (e) {}
    }
    if (hit) sh.deleteRow(i + 2);
  }
}

/* ---------- 讓人看得懂的分頁，每次存檔後自動重建 ---------- */
function owners_(s) { return String(s || '').split(/[、,，\/／]+/).map(x => x.trim()).filter(Boolean); }

function rebuildViews_() {
  const all = readAll_();
  const gname = {};
  all.group.forEach(g => { gname[g.id] = g.name; });
  const meetings = all.meeting.slice().sort((a, b) =>
    String(b.date || '').localeCompare(String(a.date || '')) || (b.createdAt || 0) - (a.createdAt || 0));
  const reps = {};
  all.report.forEach(r => { (reps[r.itemKey] = reps[r.itemKey] || []).push(r); });
  Object.keys(reps).forEach(k => reps[k].sort((a, b) => (b.at || 0) - (a.at || 0)));

  const sG = [['組別', '開會時間', '成員', '會議數']];
  all.group.forEach(g => sG.push([g.name || '', g.schedule || '', g.members || '',
    meetings.filter(m => m.groupId === g.id).length]));

  const s1 = [['組別', '會議類型', '日期', '時間', '會議名稱', '地點', '參加者', '記錄人', '議題數', '進度追蹤項目']];
  const s2 = [['組別', '會議名稱', '日期', '議題序號', '議題', '討論重點']];
  const s3 = [['組別', '會議名稱', '會議日期', '決議內容', '負責人', '類型', '原本期限', '目前狀態', '完成度', '預計完成', '最後回報']];
  const s4 = [['組別', '項目', '負責人', '回報日期', '回報人', '進度內容', '狀態', '完成度', '預計完成', '需要協助']];

  meetings.forEach(m => {
    const g = gname[m.groupId] || '未分組';
    const decs = m.decisions || [];
    s1.push([g, m.type || '', m.date || '', m.time || '', m.title || '', m.place || '', m.attendees || '',
      m.recorder || '', (m.topics || []).length, decs.filter(d => owners_(d.owner).length).length]);
    (m.topics || []).forEach((t, i) =>
      s2.push([g, m.title || '', m.date || '', i + 1, t.title || '', (t.notes || []).join('\n')]));
    decs.forEach(d => {
      const own = owners_(d.owner).join('、');
      const l = reps[m.id + ':' + d.id] || [];
      const last = l[0];
      s3.push([g, m.title || '', m.date || '', d.text || '', own, own ? '進度追蹤' : '只記錄決議', d.due || '',
        own ? (last ? last.status : '未開始') : '', own ? ((last ? last.pct || 0 : 0) + '%') : '',
        own ? ((last && last.due) || d.due || '') : '', last ? last.date || '' : '']);
      if (own) l.forEach(r => s4.push([g, d.text || '', own, r.date || '', r.who || '', r.text || '',
        r.status || '', (r.pct || 0) + '%', r.due || '', r.help || '']));
    });
  });

  writeView_('組別', sG, [140, 160, 320, 80]);
  writeView_('會議總表', s1, [110, 100, 100, 110, 200, 100, 260, 90, 70, 100]);
  writeView_('議題與討論', s2, [110, 180, 100, 70, 180, 420]);
  writeView_('決議與待辦', s3, [110, 180, 100, 260, 120, 100, 100, 110, 70, 100, 100]);
  writeView_('進度回報', s4, [110, 220, 120, 100, 90, 360, 110, 70, 100, 200]);
}

function writeView_(name, rows, widths) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.setFrozenRows(1);
  }
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.clear();
  sh.getRange(1, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#EBEBEB');
  sh.getRange(2, 1, Math.max(rows.length - 1, 1), rows[0].length).setVerticalAlignment('top').setWrap(true);
}

/* 第一次（或更新程式後）在編輯器選這個函式按「執行」 */
function setup() {
  db_();
  rebuildViews_();
  const ss = SpreadsheetApp.getActive();
  ['工作表1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
}
