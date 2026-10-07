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

// 動画を入れる共有フォルダのID（URLの /folders/ のあとの文字列）。
// サイトはこのフォルダの動画を一覧にして読み込みます。
var FOLDER_ID = "";

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

  setupWorkshop_(ss);
}

// ワークショップの満足度アンケート用のシート
var WS_Q1 = ["たのしかった", "ふつう", "つまらない"];
var WS_Q2 = ["かんたん", "ちょうどいい", "むずかしい"];
var WS_Q3 = ["やりたい", "やりたくない"];

function setupWorkshop_(ss) {
  var sheet = ss.getSheetByName("満足度") || ss.insertSheet("満足度");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["受信日時", "たのしかった？", "キットづくり", "また やりたい？", "ひとこと", "端末送信日時", "送信ID", "取消"]);
    sheet.setFrozenRows(1);
  }

  var sum = ss.getSheetByName("満足度集計") || ss.insertSheet("満足度集計");
  sum.clear();
  var R = "満足度!$H$2:$H"; // 取消の列：「取消」が付いた回答は数えない
  var rows = [["質問", "答え", "人数", "割合"]];
  var addGroup = function (title, col, items) {
    items.forEach(function (item, i) {
      var r = rows.length + 1;
      rows.push([
        i === 0 ? title : "",
        item,
        '=COUNTIFS(満足度!$' + col + '$2:$' + col + ',"' + item + '",' + R + ',"<>取消")',
        '=IFERROR(C' + r + '/SUM(C' + (rows.length + 1 - i) + ':C' + (rows.length + items.length - i) + '),0)'
      ]);
    });
  };
  addGroup("1. ワークショップはたのしかった？", "B", WS_Q1);
  addGroup("2. キットづくりは？", "C", WS_Q2);
  addGroup("3. また やりたい？", "D", WS_Q3);
  sum.getRange(1, 1, rows.length, 4).setValues(rows.map(function (r) {
    return r.map(function (v) { return typeof v === "string" && v.charAt(0) === "=" ? "" : v; });
  }));
  // 数式は setFormula で入れる（setValues だと文字として扱われる場合があるため）
  rows.forEach(function (r, i) {
    if (i === 0) { return; }
    sum.getRange(i + 1, 3).setFormula(r[2]);
    sum.getRange(i + 1, 4).setFormula(r[3]);
  });
  sum.getRange(2, 4, rows.length - 1, 1).setNumberFormat("0%");

  var total = rows.length + 2;
  sum.getRange(total, 1).setValue("回答の合計（人）");
  sum.getRange(total, 3).setFormula('=COUNTIFS(満足度!$B$2:$B,"<>",' + R + ',"<>取消")');

  sum.getRange("F1").setValue("ひとこと");
  sum.getRange("F2").setFormula(
    '=IFERROR(FILTER(満足度!E2:E, 満足度!E2:E<>"", 満足度!H2:H<>"取消"), "")'
  );

  sum.getRange("A1:D1").setFontWeight("bold").setBackground("#e8eefc");
  sum.getRange("F1").setFontWeight("bold").setBackground("#e8eefc");
  sum.getRange(total, 1, 1, 3).setFontWeight("bold");
  sum.getRange(2, 1, rows.length - 1, 1).setFontWeight("bold");
  sum.getRange(2, 3, rows.length - 1, 1).setFontWeight("bold").setFontSize(14);
  sum.setFrozenRows(1);
  sum.setColumnWidth(1, 260);
  sum.setColumnWidth(2, 140);
  sum.setColumnWidths(3, 2, 90);
  sum.setColumnWidth(5, 30);
  sum.setColumnWidth(6, 480);
  sum.getRange("F2:F").setWrap(true);
}

// 動画フォルダの中身を一覧で返す（サイトのスタート時に読み込まれる）
function doGet(e) {
  try {
    var files = DriveApp.getFolderById(FOLDER_ID).getFiles(); // ゴミ箱の中のファイルは含まれない
    var list = [];
    while (files.hasNext()) {
      var f = files.next();
      if (String(f.getMimeType()).indexOf("video/") !== 0) { continue; } // 動画以外は無視
      list.push({
        id: f.getId(),
        title: f.getName().replace(/\.[^.]+$/, ""),
        driveId: f.getId(),
        orientation: "portrait"
      });
    }
    list.sort(function (a, b) { return a.title < b.title ? -1 : a.title > b.title ? 1 : 0; });
    return json_({ ok: true, works: list });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    // まとめて送られた場合は data.records、1件だけの場合は data そのもの
    var records = Array.isArray(data.records) ? data.records : [data];

    var valid = [];     // 動画コンテストの回答
    var workshop = [];  // ワークショップの満足度
    var cancels = [];   // ワークショップの回答の取り消し
    records.forEach(function (r) {
      if (r && r.kind === "workshop") {
        if (r.op === "cancel") {
          if (r.targetId) { cancels.push(String(r.targetId).slice(0, 100)); }
          return;
        }
        if (WS_Q1.indexOf(r.q1) < 0 || WS_Q2.indexOf(r.q2) < 0 || WS_Q3.indexOf(r.q3) < 0) { return; }
        workshop.push({
          id: String(r.submissionId || "").slice(0, 100),
          q1: r.q1, q2: r.q2, q3: r.q3,
          comment: String(r.comment || "").slice(0, 500),
          sentAt: String(r.sentAt || "").slice(0, 40)
        });
        return;
      }
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
    if (valid.length === 0 && workshop.length === 0 && cancels.length === 0) {
      return json_({ ok: false, error: "invalid count" });
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var ss = getSpreadsheet_();
      var added = 0;
      var duplicates = 0;
      var now = new Date();

      if (valid.length > 0) {
        var sheet = ss.getSheetByName("回答");
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

        var rows = [];
        valid.forEach(function (r) {
          if (r.id && seen[r.id]) { duplicates++; return; }
          if (r.id) { seen[r.id] = true; }
          rows.push([now, r.workId, r.title, r.funny, r.notFunny, r.sentAt, r.id]);
        });

        // 何件でも、1回の書き込みで追加する
        if (rows.length > 0) {
          sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
        }
        added += rows.length;
      }

      if (workshop.length > 0 || cancels.length > 0) {
        var ws = ss.getSheetByName("満足度");
        if (!ws) {
          return json_({ ok: false, error: "run setup() first" });
        }
        var ids = {}; // 送信ID → 行番号
        var wlast = ws.getLastRow();
        if (wlast >= 2) {
          ws.getRange(2, 7, wlast - 1, 1).getValues().forEach(function (row, i) {
            if (row[0]) { ids[String(row[0])] = i + 2; }
          });
        }
        var wrows = [];
        workshop.forEach(function (r) {
          if (r.id && ids[r.id]) { duplicates++; return; }
          if (r.id) { ids[r.id] = wlast + 1 + wrows.length; } // 追加後の行番号
          wrows.push([now, r.q1, r.q2, r.q3, r.comment, r.sentAt, r.id, ""]);
        });
        if (wrows.length > 0) {
          ws.getRange(wlast + 1, 1, wrows.length, 8).setValues(wrows);
          added += wrows.length;
        }
        // 取り消し：行は消さず、「取消」と印を付ける（集計には数えない）
        cancels.forEach(function (id) {
          if (ids[id]) { ws.getRange(ids[id], 8).setValue("取消"); }
        });
      }

      return json_({ ok: true, added: added, duplicates: duplicates });
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
