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

  // 合計(人)：その作品の投票数（面白かった＋面白くなかった）
  summary.getRange("G1").setValue("合計(人)");
  summary.getRange("G2").setFormula('=ARRAYFORMULA(IF(C2:C="","",C2:C+D2:D))');
  // 全体の合計（I〜J列）：総投票数など
  summary.getRange("I1").setValue("全体の合計");
  summary.getRange("I2:I5").setValues([["総投票数(人)"], ["面白かった(人)"], ["面白くなかった(人)"], ["作品数"]]);
  summary.getRange("J2").setFormula('=SUM(C2:C)+SUM(D2:D)');
  summary.getRange("J3").setFormula('=SUM(C2:C)');
  summary.getRange("J4").setFormula('=SUM(D2:D)');
  summary.getRange("J5").setFormula('=COUNTA(B2:B)-COUNTIF(B2:B,"まだ回答がありません")');

  summary.getRange("A1:G1").setFontWeight("bold").setBackground("#e8eefc");
  summary.getRange("I1:J1").setFontWeight("bold").setBackground("#e8eefc");
  summary.getRange("I2:I5").setFontWeight("bold");
  summary.getRange("J2:J5").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("right");
  summary.getRange("D2:D").setFontWeight("bold");
  summary.getRange("G2:G").setFontWeight("bold").setHorizontalAlignment("right");
  summary.setFrozenRows(1);
  summary.getRange("A2:A").setHorizontalAlignment("center");
  summary.getRange("C2:C").setFontWeight("bold").setFontSize(14);
  summary.getRange("F2:F").setFontColor("#2f6fed");
  summary.setColumnWidth(1, 60);
  summary.setColumnWidth(2, 260);
  summary.setColumnWidths(3, 3, 120);
  summary.setColumnWidth(6, 360);
  summary.setColumnWidth(7, 90);
  summary.setColumnWidth(8, 30);
  summary.setColumnWidths(9, 2, 140);
  // 1位の行を目立たせる
  summary.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$A2=1")
      .setBackground("#fff2cc")
      .setBold(true)
      .setRanges([summary.getRange("A2:G200")])
      .build()
  ]);

  setupWorkshop_(ss);
}

// ワークショップの満足度アンケート用のシート
// 質問の版。質問を変えたら増やす（サイトの古い画面は「ページを更新してください」と表示される）
var WS_VERSION = 2;
var WS_Q1 = ["たのしかった", "ふつう", "たのしくなかった"];
var WS_Q2 = ["ダンボールは、リサイクルしてまたつかえる", "カプセルづくりは、しょうがいのあるひとのおしごとにつながる", "まだよくわからない"];
var WS_Q2_UNSURE = "まだよくわからない";
var WS_Q3 = ["してみたい", "まだわからない", "したくない"];
var WS_SEP = "|";   // 2番の複数回答を1つのセルに並べるときの区切り
var WS_INVALID = "無効";

// 前の版（4問）の回答。古い画面から届いた分は「満足度_旧」に記録する
var OLD_Q1 = ["たのしかった", "ふつう", "つまらない"];
var OLD_Q2 = ["かんたん", "ちょうどいい", "むずかしい"];
var OLD_Q3 = ["やりたい", "やりたくない"];
var OLD_HEADER = ["受信日時", "たのしかった？", "キットづくり", "また やりたい？", "ひとこと", "端末送信日時", "送信ID", "取消"];
var WS_HEADER = ["受信日時", "1 たのしかった？", "2 わかったこと（いくつでも）", "3 リサイクルしてみたい？", "端末送信日時", "送信ID", "取消"];

function legacySheet_(ss) {
  var sh = ss.getSheetByName("満足度_旧");
  if (!sh) {
    sh = ss.insertSheet("満足度_旧");
    sh.appendRow(OLD_HEADER);
    sh.setFrozenRows(1);
  }
  return sh;
}

function setupWorkshop_(ss) {
  // 前の版の「満足度」シートは、消さずに「満足度_旧」へ名前を変えて残す
  var sheet = ss.getSheetByName("満足度");
  if (sheet && sheet.getRange("B1").getValue() === OLD_HEADER[1] && !ss.getSheetByName("満足度_旧")) {
    sheet.setName("満足度_旧");
    sheet = null;
  }
  sheet = ss.getSheetByName("満足度") || ss.insertSheet("満足度");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(WS_HEADER);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(3, 420);
  }

  var sum = ss.getSheetByName("満足度集計") || ss.insertSheet("満足度集計");
  sum.clear();
  var G = '満足度!$G$2:$G';                 // 取消の列：「取消」が付いた回答は数えない
  var A = '満足度!$A$2:$A';                 // 受信日時（回答のある行）
  var rows = [];                             // [A, B, C(式または値), D(式または値)]
  var bold = [], titleRows = [], noteRows = [];
  var add = function (r) { rows.push(r); return rows.length; };

  var addGroup = function (title, col, items, multi) {
    var C = '満足度!$' + col + '$2:$' + col;
    var valid = 'COUNTIFS(' + C + ',"<>",' + C + ',"<>' + WS_INVALID + '",' + G + ',"<>取消")';
    var blank = 'COUNTIFS(' + A + ',"<>",' + C + ',"",' + G + ',"<>取消")';
    var bad = 'COUNTIFS(' + C + ',"' + WS_INVALID + '",' + G + ',"<>取消")';
    titleRows.push(add([title, "選択肢", "人数", "割合"]));
    var first = rows.length + 1, last = first + items.length - 1;
    var validRow = last + 1;
    items.forEach(function (item) {
      var r = rows.length + 1;
      var crit = multi ? '"*"&"' + item + '"&"*"' : '"' + item + '"';
      add(["", item, '=COUNTIFS(' + C + ',' + crit + ',' + G + ',"<>取消")',
           '=IFERROR(C' + r + '/$C$' + validRow + ',0)']);
    });
    // 有効回答の件数（割合のもと）。1行にまとめて「有効回答：○件　未回答：○件　無効回答：○件」と表示する
    add(["", "有効回答：", '=' + valid, ""]);
    var lineRow = add(["", '="有効回答："&' + valid + '&"件　未回答："&' + blank + '&"件　無効回答："&' + bad + '&"件"', "", ""]);
    noteRows.push(lineRow);
    if (multi) { add(["", "※いくつでも えらべるため、割合の合計は100%をこえることがあります", "", ""]); }
    add(["", "", "", ""]);
  };
  addGroup("1 ペパポンをつくって、たのしかった", "B", WS_Q1, false);
  addGroup("2 きょう、わかったことはどれかな（いくつでも）", "C", WS_Q2, true);
  addGroup("3 おうちのひとと、リサイクルをしてみたい", "D", WS_Q3, false);
  var total = add(["回答の合計（人）", "", '=COUNTIFS(' + A + ',"<>",' + G + ',"<>取消")', ""]);

  rows.forEach(function (r, i) {
    var row = i + 1;
    sum.getRange(row, 1).setValue(r[0]);
    [1, 2, 3].forEach(function (k) {
      var v = r[k];
      if (v === "") { return; }
      var cell = sum.getRange(row, k + 1);
      if (typeof v === "string" && v.charAt(0) === "=") { cell.setFormula(v); } else { cell.setValue(v); }
    });
    if (r[3] !== "" && r[3] !== undefined && String(r[3]).indexOf("IFERROR") >= 0) {
      sum.getRange(row, 4).setNumberFormat("0%");
    }
  });
  titleRows.forEach(function (r) {
    sum.getRange(r, 1, 1, 4).setFontWeight("bold").setBackground("#e8eefc");
  });
  noteRows.forEach(function (r) { sum.getRange(r, 2).setFontWeight("bold"); });
  sum.getRange(total, 1, 1, 3).setFontWeight("bold");
  sum.getRange("C1:C" + total).setFontWeight("bold").setFontSize(14);
  sum.setColumnWidth(1, 330);
  sum.setColumnWidth(2, 420);
  sum.setColumnWidths(3, 2, 90);
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
    return json_({ ok: true, works: list, wsVersion: WS_VERSION });
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
    var workshop = [];  // ワークショップの満足度（今の版）
    var legacy = [];    // ワークショップの満足度（前の版）
    var cancels = [];   // ワークショップの回答の取り消し
    records.forEach(function (r) {
      if (r && r.kind === "workshop") {
        if (r.op === "cancel") {
          if (r.targetId) { cancels.push(String(r.targetId).slice(0, 100)); }
          return;
        }
        var wid = String(r.submissionId || "").slice(0, 100);
        var wsent = String(r.sentAt || "").slice(0, 40);
        if (Number(r.v) === WS_VERSION) {
          // 今の版：空は「未回答」、決まった選択肢以外や矛盾した組み合わせは「無効」
          var one = function (v, list) {
            if (v === undefined || v === null || v === "") { return ""; }
            return list.indexOf(v) >= 0 ? v : WS_INVALID;
          };
          var a1 = one(r.q1, WS_Q1), a3 = one(r.q3, WS_Q3), a2 = "";
          var picked = Array.isArray(r.q2) ? r.q2 : (r.q2 ? [r.q2] : []);
          if (picked.length > 0) {
            var okAll = picked.every(function (v) { return WS_Q2.indexOf(v) >= 0; });
            var uniq = WS_Q2.filter(function (v) { return picked.indexOf(v) >= 0; });
            var clash = uniq.indexOf(WS_Q2_UNSURE) >= 0 && uniq.length > 1;
            a2 = (!okAll || clash) ? WS_INVALID : uniq.join(WS_SEP);
          }
          if (a1 === "" && a2 === "" && a3 === "") { return; }
          workshop.push({ id: wid, q1: a1, q2: a2, q3: a3, sentAt: wsent });
        } else if (r.v === undefined &&
                   OLD_Q1.indexOf(r.q1) >= 0 && OLD_Q2.indexOf(r.q2) >= 0 && OLD_Q3.indexOf(r.q3) >= 0) {
          // 前の版の画面から届いた回答は「満足度_旧」へ
          legacy.push({ id: wid, q1: r.q1, q2: r.q2, q3: r.q3,
                        comment: String(r.comment || "").slice(0, 500), sentAt: wsent });
        }
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
    if (valid.length === 0 && workshop.length === 0 && legacy.length === 0 && cancels.length === 0) {
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

      if (workshop.length > 0 || legacy.length > 0 || cancels.length > 0) {
        var ws = ss.getSheetByName("満足度");
        if (!ws) {
          return json_({ ok: false, error: "run setup() first" });
        }
        var ids = {};    // 今の版：送信ID → 行番号（G列に取消）
        var wlast = ws.getLastRow();
        if (wlast >= 2) {
          ws.getRange(2, 6, wlast - 1, 1).getValues().forEach(function (row, i) {
            if (row[0]) { ids[String(row[0])] = i + 2; }
          });
        }
        var wrows = [];
        workshop.forEach(function (r) {
          if (r.id && ids[r.id]) { duplicates++; return; }
          if (r.id) { ids[r.id] = wlast + 1 + wrows.length; } // 追加後の行番号
          wrows.push([now, r.q1, r.q2, r.q3, r.sentAt, r.id, ""]);
        });
        if (wrows.length > 0) {
          ws.getRange(wlast + 1, 1, wrows.length, 7).setValues(wrows);
          added += wrows.length;
        }

        // 前の版（満足度_旧）：記録と取り消し
        var lids = {};   // 送信ID → 行番号（H列に取消）
        var lsh = (legacy.length > 0) ? legacySheet_(ss) : ss.getSheetByName("満足度_旧");
        if (lsh) {
          var llast = lsh.getLastRow();
          if (llast >= 2) {
            lsh.getRange(2, 7, llast - 1, 1).getValues().forEach(function (row, i) {
              if (row[0]) { lids[String(row[0])] = i + 2; }
            });
          }
          var lrows = [];
          legacy.forEach(function (r) {
            if (r.id && lids[r.id]) { duplicates++; return; }
            if (r.id) { lids[r.id] = llast + 1 + lrows.length; }
            lrows.push([now, r.q1, r.q2, r.q3, r.comment, r.sentAt, r.id, ""]);
          });
          if (lrows.length > 0) {
            lsh.getRange(llast + 1, 1, lrows.length, 8).setValues(lrows);
            added += lrows.length;
          }
        }

        // 取り消し：行は消さず、「取消」と印を付ける（集計には数えない）
        cancels.forEach(function (id) {
          if (ids[id]) { ws.getRange(ids[id], 7).setValue("取消"); }
          else if (lids[id]) { lsh.getRange(lids[id], 8).setValue("取消"); }
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
