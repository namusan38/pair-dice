/**
 * ペア・ダイス ランキング用 Google Apps Script
 *
 * スプレッドシート「ペア・ダイス ランキング」の
 * 拡張機能 → Apps Script に貼り付けて、ウェブアプリとしてデプロイします。
 * 1行目は見出し(日時 / 名前 / スコア / 手番)。
 */

const TOP_N = 20;          // ランキングに表示する件数
const MAX_NAME = 12;       // 名前の最大文字数
const MIN_SCORE = -2200;   // 理論上の最低点
const MAX_SCORE = 4020;    // 理論上の最高点

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'submit') return json(submit_(p));
    return json({ rows: top_() });
  } catch (err) {
    return json({ error: String(err.message || err) });
  }
}

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function top_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, 4).getValues();
  return values
    .filter(r => r[1] !== '' && r[2] !== '')
    .map(r => ({ name: String(r[1]).replace(/^'/, ''), score: Number(r[2]), turns: Number(r[3]), at: r[0] }))
    .sort((a, b) => b.score - a.score || a.turns - b.turns)
    .slice(0, TOP_N)
    .map(r => ({ name: r.name, score: r.score, turns: r.turns }));
}

function submit_(p) {
  let name = String(p.name || '').trim().slice(0, MAX_NAME);
  const score = Number(p.score);
  const turns = Number(p.turns);
  if (!name) throw new Error('名前がありません');
  if (!Number.isInteger(score) || score < MIN_SCORE || score > MAX_SCORE) throw new Error('スコアが不正です');
  if (!Number.isInteger(turns) || turns < 1 || turns > 200) throw new Error('手番が不正です');
  // 数式として解釈されないようにする
  if (/^[=+\-@]/.test(name)) name = "'" + name;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet_().appendRow([new Date(), name, score, turns]);
  } finally {
    lock.releaseLock();
  }
  const all = top_();
  const idx = all.findIndex(r => r.score === score && r.turns === turns && r.name === name.replace(/^'/, ''));
  return { ok: true, rank: idx >= 0 ? idx + 1 : null };
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
