/**
 * 学習の記録を受けとって、スプレッドシートにためるプログラム（Google Apps Script）
 *
 * 使い方はこのフォルダの「セットアップ.md」を見てください。
 * ここを直したら、かならず「デプロイ → デプロイを管理 → 編集 → 新バージョン」で
 * 公開しなおしてください（保存しただけでは反映されません）。
 */

// ★ここだけ、自分で決めた合言葉に書きかえてください（iPad と保護者のスマホに入れるものと同じ）
const KEY = 'あいことば';

// ---- 受けとる（ドリルからの送信） ----
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (String(data.k || '') !== KEY) return text('NG:key');
    const name = String(data.name || '').slice(0, 20);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();

    const ses = [], items = [], totals = [];
    for (const r of rows) {
      if (r[0] === 's') ses.push([r[2], r[1], r[3], r[4], r[5], r[6], r[7], name, now]);
      else if (r[0] === 'i') items.push([r[1], r[2], name, r[3], r[4], r[5], r[6], r[7], r[8], r[9] || '', now]);
      else if (r[0] === 't') totals.push([r[1], r[2], now]);
    }
    if (ses.length) {
      const sh = sheet(ss, 'sessions', ['日', '教科', '分野', '復習', '出題数', '正解数', '秒', '名前', '受信']);
      sh.getRange(sh.getLastRow() + 1, 1, ses.length, ses[0].length).setValues(ses);
    }
    if (items.length) upsertItems(ss, items);
    if (totals.length) upsertTotals(ss, totals);
    return text('OK');
  } catch (err) {
    return text('NG:' + err);
  }
}

// ---- 読みだす（保護者ページ oya.html から） ----
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (String(p.k || '') !== KEY) return out({ ok: false, error: 'key' }, p.callback);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ses = values(ss, 'sessions').map(function (r) {
    return [num(r[0]), str(r[1]), str(r[2]), num(r[3]), num(r[4]), num(r[5]), num(r[6]), str(r[7])];
  });
  const items = values(ss, 'items').map(function (r) {
    return [str(r[0]), str(r[1]), str(r[3]), num(r[4]), num(r[5]), num(r[6]), num(r[7]), num(r[8]), str(r[9])];
  });
  const total = {};
  values(ss, 'totals').forEach(function (r) { total[str(r[0])] = num(r[1]); });
  // 多くなりすぎたら、新しいものだけ返す
  const cut = Math.max(0, ses.length - 4000);
  return out({ ok: true, ses: ses.slice(cut), items: items, total: total, at: new Date().toISOString() }, p.callback);
}

// ---- ここから下は、しくみの部分（さわらなくて大丈夫）----
function sheet(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function upsertItems(ss, items) {
  const sh = sheet(ss, 'items', ['教科', '項目', '名前', 'ラベル', '出題数', '正解数', 'まちがい', '最後の日', '卒業日', '分野', '受信']);
  const last = sh.getLastRow();
  const cur = last > 1 ? sh.getRange(2, 1, last - 1, 3).getValues() : [];
  const idx = {};
  cur.forEach(function (r, i) { idx[r[0] + '\t' + r[1] + '\t' + r[2]] = i + 2; });
  const add = [];
  for (const it of items) {
    const row = idx[it[0] + '\t' + it[1] + '\t' + it[2]];
    if (row) sh.getRange(row, 1, 1, it.length).setValues([it]);
    else add.push(it);
  }
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, add[0].length).setValues(add);
}
function upsertTotals(ss, totals) {
  const sh = sheet(ss, 'totals', ['教科', '項目の総数', '受信']);
  const last = sh.getLastRow();
  const cur = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  const idx = {};
  cur.forEach(function (r, i) { idx[r[0]] = i + 2; });
  for (const t of totals) {
    const row = idx[t[0]];
    if (row) sh.getRange(row, 1, 1, t.length).setValues([t]);
    else sh.getRange(sh.getLastRow() + 1, 1, 1, t.length).setValues([t]);
  }
}
function values(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}
function num(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function str(v) { return v === null || v === undefined ? '' : String(v); }
function text(s) { return ContentService.createTextOutput(s); }
function out(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
