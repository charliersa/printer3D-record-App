/**
 * 列印工作台 — Google Sheet 後端
 *
 * 部署方式看同資料夾的 README.md。重點只有三個：
 *   1. 把下面的 TOKEN 改成自己的長字串，前端要填一樣的。
 *   2. 部署 → 新增部署作業 → 網頁應用程式，「執行身分」選自己、「誰可以存取」選任何人。
 *   3. 把 /exec 結尾的網址貼到 App 的「設定 → 雲端同步」。
 *
 * 前端一律用 POST + Content-Type: text/plain 呼叫。
 * 這是刻意的：Apps Script 不會回應 CORS preflight（OPTIONS），
 * 送 application/json 會被瀏覽器擋掉，text/plain 才算「簡單請求」。
 */

/** 前端要送一模一樣的字串才會被受理。務必改掉。 */
const TOKEN = 'CHANGE-ME-換成你自己的長字串';

/** 留空 = 用這份指令碼所綁定的試算表。要寫到別份就填試算表 ID。 */
const SPREADSHEET_ID = '';

const META_SHEET = '_meta';
const SCHEMA_SHEET = '_schema';
const SCHEMA_HEADER = ['id', 'sheet', 'code', 'name', 'tint', 'builtin', 'listFields', 'fields'];

/* ------------------------------------------------------------------ *
 * 進入點
 * ------------------------------------------------------------------ */

/** 用瀏覽器直接開 /exec 會看到這個，方便確認部署成功。 */
function doGet() {
  return json({ ok: true, service: '列印工作台 backend', time: new Date().toISOString() });
}

function doPost(e) {
  try {
    if (TOKEN.indexOf('CHANGE-ME') === 0) {
      return json({ ok: false, code: 'no_token', error: '後端的 TOKEN 還沒改，請先改掉再重新部署。' });
    }
    const body = parseBody_(e);
    if (!body) return json({ ok: false, code: 'bad_request', error: '沒有收到 JSON。' });
    if (body.token !== TOKEN) return json({ ok: false, code: 'unauthorized', error: '密鑰不對。' });

    switch (body.action) {
      case 'ping': return json(ping_());
      case 'pull': return json(pull_());
      case 'push': return json(push_(body));
      default: return json({ ok: false, code: 'bad_action', error: '不認得的 action：' + body.action });
    }
  } catch (err) {
    return json({ ok: false, code: 'error', error: String((err && err.message) || err) });
  }
}

/* ------------------------------------------------------------------ *
 * 三個動作
 * ------------------------------------------------------------------ */

function ping_() {
  const ss = book_();
  const meta = readMeta_(ss);
  return {
    ok: true,
    title: ss.getName(),
    sheetUrl: ss.getUrl(),
    rev: meta.rev,
    updatedAt: meta.updatedAt,
    collections: readSchema_(ss).length
  };
}

/** 讀回整包資料，格式跟 App 的 localStorage 一樣。 */
function pull_() {
  const ss = book_();
  const meta = readMeta_(ss);
  const schema = readSchema_(ss);
  const tz = ss.getSpreadsheetTimeZone();

  const collections = schema.map(function (row) {
    const fields = row.fields;
    const items = [];
    const sh = ss.getSheetByName(row.sheet);
    if (sh) {
      const lastRow = sh.getLastRow();
      const width = Math.min(fields.length + 1, Math.max(sh.getLastColumn(), 1));
      if (lastRow > 1 && width > 0) {
        sh.getRange(2, 1, lastRow - 1, width).getValues().forEach(function (r) {
          const blank = r.every(function (v) { return v === '' || v === null; });
          if (blank) return;                       // 使用者在試算表按 Enter 留下的空列
          const item = { id: String(r[0] || '').trim() || ('s' + Utilities.getUuid().slice(0, 8)) };
          fields.forEach(function (f, i) { item[f.key] = cellIn_(r[i + 1], f.type, tz); });
          items.push(item);
        });
      }
    }
    return {
      id: row.id, code: row.code, name: row.name, tint: row.tint,
      builtin: row.builtin, fields: fields, listFields: row.listFields, items: items
    };
  });

  return { ok: true, rev: meta.rev, updatedAt: meta.updatedAt, colId: meta.colId, collections: collections };
}

/**
 * 整包覆蓋雲端。帶 baseRev 時會擋掉「雲端比較新」的情況，
 * 除非前端明講 force（使用者按了強制覆蓋）。
 */
function push_(body) {
  const cols = body.collections;
  if (!Array.isArray(cols) || !cols.length) throw new Error('collections 是空的，不覆蓋。');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('另一個同步還在進行，稍後再試。');
  try {
    const ss = book_();
    const cur = readMeta_(ss);
    if (!body.force && body.baseRev != null && cur.rev && Number(body.baseRev) !== Number(cur.rev)) {
      return {
        ok: false, code: 'conflict', rev: cur.rev, updatedAt: cur.updatedAt,
        error: '雲端資料比你手上的新（雲端 rev ' + cur.rev + '、你的 ' + body.baseRev + '）。'
      };
    }

    const prev = {};
    readSchema_(ss).forEach(function (r) { prev[r.id] = r.sheet; });

    const names = {};
    cols.forEach(function (c) { names[c.id] = writeCollection_(ss, c, prev[c.id]); });

    // 只清掉「這次沒送上來的分類」留下的工作表，使用者自己加的分頁不動
    Object.keys(prev).forEach(function (id) {
      if (names[id]) return;
      const sh = ss.getSheetByName(prev[id]);
      if (sh && ss.getSheets().length > 1) ss.deleteSheet(sh);
    });

    writeSchema_(ss, cols, names);

    const rev = (Number(cur.rev) || 0) + 1;
    const updatedAt = new Date().toISOString();
    writeMeta_(ss, { rev: rev, updatedAt: updatedAt, colId: body.colId || cols[0].id });
    return { ok: true, rev: rev, updatedAt: updatedAt, collections: cols.length };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ *
 * 工作表讀寫
 * ------------------------------------------------------------------ */

/** 一個分類一張工作表：第一列是欄位名稱，A 欄是 App 內部的 id。 */
function writeCollection_(ss, c, prevName) {
  const want = safeSheetName_(c.name || c.id);
  let sh = prevName ? ss.getSheetByName(prevName) : null;
  if (!sh) sh = ss.getSheetByName(want);
  if (!sh) sh = ss.insertSheet(want);
  else if (sh.getName() !== want && !ss.getSheetByName(want)) sh.setName(want);

  const fields = Array.isArray(c.fields) ? c.fields : [];
  const header = ['id'].concat(fields.map(function (f) { return f.label || f.key; }));
  const rows = (c.items || []).map(function (i) {
    return [i.id].concat(fields.map(function (f) { return cellOut_(i[f.key], f.type); }));
  });

  sh.clear();
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
  try { sh.autoResizeColumns(1, header.length); } catch (e) { /* 欄位很多時偶爾會逾時，不重要 */ }
  return sh.getName();
}

function writeSchema_(ss, cols, names) {
  const sh = sheet_(ss, SCHEMA_SHEET);
  const rows = cols.map(function (c) {
    return [
      c.id,
      names[c.id] || safeSheetName_(c.name || c.id),
      c.code || '',
      c.name || '',
      c.tint || '',
      c.builtin ? 'TRUE' : 'FALSE',
      Array.isArray(c.listFields) ? JSON.stringify(c.listFields) : '',
      JSON.stringify(c.fields || [])
    ];
  });
  sh.clear();
  sh.getRange(1, 1, 1, SCHEMA_HEADER.length).setValues([SCHEMA_HEADER]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, SCHEMA_HEADER.length).setValues(rows);
  sh.setFrozenRows(1);
  hide_(ss, sh);
}

function readSchema_(ss) {
  const sh = ss.getSheetByName(SCHEMA_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, SCHEMA_HEADER.length).getValues()
    .filter(function (r) { return String(r[0] || '').trim(); })
    .map(function (r) {
      return {
        id: String(r[0]).trim(),
        sheet: String(r[1] || '').trim() || String(r[3] || '').trim(),
        code: String(r[2] || ''),
        name: String(r[3] || ''),
        tint: String(r[4] || ''),
        builtin: String(r[5]).toUpperCase() === 'TRUE',
        listFields: parseJson_(r[6], null),
        fields: parseJson_(r[7], [])
      };
    });
}

function readMeta_(ss) {
  const sh = ss.getSheetByName(META_SHEET);
  const out = { rev: 0, updatedAt: '', colId: '' };
  if (!sh || sh.getLastRow() < 1) return out;
  sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(function (r) {
    const k = String(r[0] || '').trim();
    if (k) out[k] = r[1];
  });
  out.rev = Number(out.rev) || 0;
  out.updatedAt = out.updatedAt ? String(out.updatedAt) : '';
  out.colId = out.colId ? String(out.colId) : '';
  return out;
}

function writeMeta_(ss, meta) {
  const sh = sheet_(ss, META_SHEET);
  sh.clear();
  sh.getRange(1, 1, 3, 2).setValues([
    ['rev', meta.rev],
    ['updatedAt', meta.updatedAt],
    ['colId', meta.colId]
  ]);
  hide_(ss, sh);
}

/* ------------------------------------------------------------------ *
 * 小工具
 * ------------------------------------------------------------------ */

function book_() {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('找不到試算表。這份指令碼要綁在試算表上，或是填 SPREADSHEET_ID。');
  return ss;
}

function sheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function hide_(ss, sh) {
  // 只剩一張可見工作表時 hideSheet() 會丟錯，忽略即可
  try { if (ss.getSheets().filter(function (s) { return !s.isSheetHidden(); }).length > 1) sh.hideSheet(); } catch (e) {}
}

function safeSheetName_(name) {
  const s = String(name || '').replace(/[\[\]\*\?\/\\:]/g, ' ').trim().slice(0, 90);
  return s || 'sheet';
}

function parseBody_(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (err) { return null; }
  }
  if (e && e.parameter && e.parameter.payload) {
    try { return JSON.parse(e.parameter.payload); } catch (err) { return null; }
  }
  return null;
}

function parseJson_(v, dflt) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return dflt;
  try { return JSON.parse(s); } catch (e) { return dflt; }
}

/** App → 試算表。數字欄位存成真數字，這樣在試算表裡才能加總。 */
function cellOut_(v, type) {
  if (v === undefined || v === null || v === '') return '';
  if (type === 'number') {
    const n = Number(v);
    return isNaN(n) ? String(v) : n;
  }
  const s = String(v);
  // 開頭是 = 會被當成公式，加單引號讓它維持文字（讀回來時不會帶著單引號）
  return s.charAt(0) === '=' ? "'" + s : s;
}

/** 試算表 → App。 */
function cellIn_(v, type, tz) {
  if (v === undefined || v === null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz || 'Asia/Taipei', type === 'date' ? 'yyyy-MM-dd' : 'yyyy-MM-dd HH:mm');
  }
  if (type === 'number') {
    const n = Number(v);
    return isNaN(n) ? String(v) : n;
  }
  return String(v);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ *
 * 手動跑一次就好：建立 _meta / _schema，並在紀錄裡印出試算表網址
 * ------------------------------------------------------------------ */
function setup() {
  const ss = book_();
  sheet_(ss, META_SHEET);
  sheet_(ss, SCHEMA_SHEET);
  if (readMeta_(ss).rev === 0) writeMeta_(ss, { rev: 0, updatedAt: '', colId: '' });
  Logger.log('試算表：%s\n%s', ss.getName(), ss.getUrl());
  Logger.log('接著到「部署 → 新增部署作業 → 網頁應用程式」取得 /exec 網址。');
}
