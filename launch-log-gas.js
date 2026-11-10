// 麻雀スタッツ 起動ログ（Google Apps Script）
// 使い方：スプレッドシートの「拡張機能 → Apps Script」に貼り付け →「デプロイ → 新しいデプロイ → ウェブアプリ」
//         実行ユーザー：自分／アクセスできるユーザー：全員 → 発行されたURLを各HTMLの LOG_URL に入れる
// シート「log」が無ければ自動で作り、見出し行を入れる。

const TOKEN = 'masu-2026';          // 各HTMLの LOG_TOKEN と同じ文字列にする
const SHEET_NAME = 'log';
const APPS = { input: 'スマホ入力', viewer: '分析ビューア', index: 'トップ' };

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out('ng'); }
  if (!d || d.token !== TOKEN || !APPS[d.app]) return out('ng');

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(['日時', 'アプリ', '端末ID', '起動方法', '端末情報']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([
      new Date(),
      APPS[d.app],
      String(d.device || '').slice(0, 16),
      String(d.mode || '').slice(0, 10),
      String(d.ua || '').slice(0, 300),
    ]);
  } finally {
    lock.releaseLock();
  }
  return out('ok');
}

function out(s) {
  return ContentService.createTextOutput(s);
}
