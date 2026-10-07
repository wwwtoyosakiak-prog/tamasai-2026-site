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
    answers.appendRow(["受信日時", "作品ID", "作品名", "面白かった(人)", "面白くなかった(人)", "端末送信日時", "送信ID"]);
    answers.setFrozenRows(1);
  }

  // 集計シート：「面白かった」の人数が多い順に並べ、順位と棒を付ける
  var summary = ss.getSheetByName("集計") || ss.insertSheet("集計");
  summary.clear();
  summary.clearConditionalFormatRules();

  summary.getRange("A1").setValue("順位");
  // B〜E列：作品ごとの合計。面白かった(人)が多い順、同数なら面白くなかった(人)が少ない順
  summary.getRange("B1").setFormula(
    '=IFERROR(QUERY(回答!A2:F, "select C, sum(D), sum(E), count(C) where C is not null ' +
    "group by C order by sum(D) desc, sum(E) asc " +
    "label C '作品名', sum(D) '面白かった(人)', sum(E) '面白くなかった(人)', count(C) '回答数'\", 0), " +
    '"まだ回答がありません")'
  );
  // 順位：面白かった(人)が多いほど上位。同数なら同じ順位
  summary.getRange("A2").setFormula('=ARRAYFORMULA(IF(C2:C="","",1+COUNTIF(C2:C,">"&C2:C)))');
  // 棒：面白かった(人)の数だけ ■ を並べる
  summary.getRange("F1").setValue("面白かった（棒）");
  summary.getRange("F2").setFormula('=ARRAYFORMULA(IF(C2:C="","",REPT("■",C2:C)))');

  summary.getRange("A1:F1").setFontWeight("bold").setBackground("#e8eefc");
  summary.setFrozenRows(1);
  summary.getRange("A2:A").setHorizontalAlignment("center");
  summary.getRange("C2:C").setFontWeight("bold").setFontSize(14);
  summary.getRange("F2:F").setFontColor("#2f6fed");
  summary.setColumnWidth(1, 60);
  summary.setColumnWidth(2, 260);
  summary.setColumnWidths(3, 3, 120);
  summary.setColumnWidth(6, 360);
  // 1位の行を目立たせる
  summary.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$A2=1")
      .setBackground("#fff2cc")
      .setBold(true)
      .setRanges([summary.getRange("A2:F200")])
      .build()
  ]);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    // まとめて送られた場合は data.records、1件だけの場合は data そのもの
    var records = Array.isArray(data.records) ? data.records : [data];

    var valid = [];
    records.forEach(function (r) {
      var funny = toCount_(r.funny);
      var notFunny = toCount_(r.notFunny);
      if (funny === null || notFunny === null || funny + notFunny < 1) { return; }
      valid.push({
        id: String(r.submissionId || "").slice(0, 100),
        workId: String(r.workId || "").slice(0, 100),
        title: String(r.title || "").slice(0, 200),
        funny: funny,
        notFunny: notFunny,
        sentAt: String(r.sentAt || "").slice(0, 40)
      });
    });
    if (valid.length === 0) {
      return json_({ ok: false, error: "invalid count" });
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var sheet = getSpreadsheet_().getSheetByName("回答");
      if (!sheet) {
        return json_({ ok: false, error: "run setup() first" });
      }
      if (!sheet.getRange("G1").getValue()) { sheet.getRange("G1").setValue("送信ID"); }

      // 再送で同じ回答が二重に入らないよう、記録済みの送信IDを1回でまとめて読む
      var seen = {};
      var last = sheet.getLastRow();
      if (last >= 2) {
        sheet.getRange(2, 7, last - 1, 1).getValues().forEach(function (row) {
          if (row[0]) { seen[String(row[0])] = true; }
        });
      }

      var now = new Date();
      var rows = [];
      var duplicates = 0;
      valid.forEach(function (r) {
        if (r.id && seen[r.id]) { duplicates++; return; }
        if (r.id) { seen[r.id] = true; }
        rows.push([now, r.workId, r.title, r.funny, r.notFunny, r.sentAt, r.id]);
      });

      // 何件でも、1回の書き込みで追加する
      if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
      }
      return json_({ ok: true, added: rows.length, duplicates: duplicates });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
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
