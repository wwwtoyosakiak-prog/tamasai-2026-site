/**
 * 多摩祭2026 動画コンテスト：アンケート記録用スクリプト（Google Apps Script）
 *
 * 使い方は README.md の「アンケートの記録先を設定する」を参照。
 * 1. setup() を1回だけ実行する（シート「回答」「集計」を作ります）
 * 2. 「デプロイ」→「ウェブアプリ」として公開し、URL を config.js に貼る
 */

var MAX_PEOPLE = 10;

// 記録先スプレッドシートのID（URLの /d/ と /edit の間の文字列）。
// スプレッドシートの「拡張機能 → Apps Script」から作った場合は、空のままで構いません。
var SHEET_ID = "";

function getSpreadsheet_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function setup() {
  var ss = getSpreadsheet_();

  var answers = ss.getSheetByName("回答") || ss.insertSheet("回答");
  if (answers.getLastRow() === 0) {
    answers.appendRow(["受信日時", "作品ID", "作品名", "面白かった(人)", "面白くなかった(人)", "端末送信日時"]);
    answers.setFrozenRows(1);
  }

  var summary = ss.getSheetByName("集計") || ss.insertSheet("集計");
  summary.clear();
  summary.getRange("A1").setFormula(
    '=IFERROR(QUERY(回答!A2:F, "select C, sum(D), sum(E), count(C) group by C ' +
    "label C '作品名', sum(D) '面白かった(人)', sum(E) '面白くなかった(人)', count(C) '回答数'\", 0), \"まだ回答がありません\")"
  );
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var data = JSON.parse(e.postData.contents);
    var funny = toCount_(data.funny);
    var notFunny = toCount_(data.notFunny);
    if (funny === null || notFunny === null || funny + notFunny < 1) {
      return json_({ ok: false, error: "invalid count" });
    }

    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName("回答");
    if (!sheet) {
      return json_({ ok: false, error: "run setup() first" });
    }

    sheet.appendRow([
      new Date(),
      String(data.workId || "").slice(0, 100),
      String(data.title || "").slice(0, 200),
      funny,
      notFunny,
      String(data.sentAt || "").slice(0, 40)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function toCount_(v) {
  var n = Number(v);
  if (!isFinite(n) || n % 1 !== 0 || n < 0 || n > MAX_PEOPLE) { return null; }
  return n;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
