/**
 * 會議記錄系統：Google 試算表資料庫
 * 貼到這份試算表的「擴充功能 → Apps Script」，再部署成「網頁應用程式」。
 */
const DATA_SHEET = '_data';

function doGet(e) {
  try {
    return json_({ ok: true, meetings: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'save') save_(req.meeting);
    else if (req.action === 'delete') del_(req.id);
    else return json_({ ok: false, error: 'unknown action' });
    rebuildViews_();
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function dataSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(1, 1, 1, 3).setValues([['id', 'updatedAt', 'json']]);
    sh.hideSheet();
  }
  return sh;
}

function rows_() {
  const sh = dataSheet_();
  const n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, 3).getValues();
}

function readAll_() {
  return rows_()
    .filter(r => r[0])
    .map(r => { try { return JSON.parse(r[2]); } catch (e) { return null; } })
    .filter(Boolean);
}

function save_(m) {
  if (!m || !m.id) throw new Error('missing id');
  const text = JSON.stringify(m);
  if (text.length > 49000) throw new Error('這場會議內容太長，請分成兩場記錄');
  const sh = dataSheet_();
  const ids = rows_().map(r => String(r[0]));
  const i = ids.indexOf(String(m.id));
  const row = [String(m.id), Number(m.updatedAt) || Date.now(), text];
  if (i >= 0) {
    const stored = Number(sh.getRange(i + 2, 2).getValue()) || 0;
    if (row[1] < stored) return; // 比較舊的版本，不覆蓋
    sh.getRange(i + 2, 1, 1, 3).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}

function del_(id) {
  const sh = dataSheet_();
  const ids = rows_().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  if (i >= 0) sh.deleteRow(i + 2);
}

/* 讓人看得懂的三個分頁，每次存檔後自動重建 */
function isTodo_(d) { return d && String(d.owner || '').trim() !== ''; }
function meetingNo_(m) {
  return 'M-' + String(m.date || '').replace(/-/g, '').slice(2) + '-' + String(m.id).slice(-4);
}

function rebuildViews_() {
  const list = readAll_().sort((a, b) =>
    String(b.date || '').localeCompare(String(a.date || '')) || (b.createdAt || 0) - (a.createdAt || 0));
  const s1 = [['會議編號', '日期', '時間', '會議名稱', '地點', '參加者', '記錄人', '範本', '議題數', '未完成待辦']];
  const s2 = [['會議編號', '會議名稱', '議題序號', '議題', '狀態', '討論重點']];
  const s3 = [['會議編號', '會議名稱', '決議內容', '時間安排', '類型', '負責人', '狀態']];
  list.forEach(m => {
    const no = meetingNo_(m);
    const td = (m.decisions || []).filter(isTodo_);
    s1.push([no, m.date || '', m.time || '', m.title || '', m.place || '', m.attendees || '', m.recorder || '',
      m.template || '', (m.topics || []).length, td.filter(d => d.status !== '已完成').length]);
    (m.topics || []).forEach((t, i) =>
      s2.push([no, m.title || '', i + 1, t.title || '', t.status || '未討論', (t.notes || []).join('\n')]));
    (m.decisions || []).forEach(d =>
      s3.push([no, m.title || '', d.text || '', d.due || '', isTodo_(d) ? '待辦' : '決議', d.owner || '',
        isTodo_(d) ? (d.status || '未開始') : '']));
  });
  writeView_('會議總表', s1, [130, 100, 110, 200, 100, 260, 90, 120, 70, 90]);
  writeView_('議題與討論', s2, [130, 180, 70, 180, 90, 420]);
  writeView_('決議與待辦', s3, [130, 180, 280, 100, 70, 90, 90]);
}

function writeView_(name, rows, widths) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
  sh.clearContents();
  sh.getRange(1, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#EBEBEB');
  sh.getRange(2, 1, Math.max(rows.length - 1, 1), rows[0].length).setVerticalAlignment('top').setWrap(true);
}

/* 第一次授權用：在編輯器選這個函式按「執行」 */
function setup() {
  dataSheet_();
  rebuildViews_();
  const first = SpreadsheetApp.getActive().getSheetByName('工作表1') || SpreadsheetApp.getActive().getSheetByName('Sheet1');
  if (first && SpreadsheetApp.getActive().getSheets().length > 1) SpreadsheetApp.getActive().deleteSheet(first);
}
