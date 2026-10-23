/**
 * 多摩祭2026 動画コンテスト：アンケート記録用スクリプト（Google Apps Script）
 *
 * 使い方は README.md の「アンケートの記録先を設定する」を参照。
 * 1. setup() を実行する（シートを作り、集計・見た目・並び順を整えます。何回実行しても記録は消えません）
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

// サイトのURL・動画フォルダ（「使い方」シートに表示する）
var SITE_URL = "https://wwwtoyosakiak-prog.github.io/tamasai-2026-site/";

var HEAD_BG = "#e8eefc";     // 見出しの色
var VIEW_COLOR = "#2f6fed";  // 見る用シートのタブの色（青）
var DATA_COLOR = "#9aa0a6";  // 記録用シートのタブの色（灰）
var HELP_COLOR = "#34a853";  // 使い方シートのタブの色（緑）

function setup() {
  var ss = getSpreadsheet_();
  var help = setupHelp_(ss);
  var answers = setupVideoAnswers_(ss);
  var summary = setupVideoSummary_(ss);
  var ws = setupWorkshop_(ss);
  var wsSum = ss.getSheetByName("満足度集計");

  // 何も入っていない「シート1」と、記録の無い「満足度_旧」は消す
  removeIfEmpty_(ss, "シート1", 0);
  removeIfEmpty_(ss, "満足度_旧", 1);

  // 並び順：使い方 → 集計（動画）→ 満足度集計（ワークショップ）→ 回答 → 満足度
  [help, summary, wsSum, answers, ws].forEach(function (sh, i) {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  });
  help.setTabColor(HELP_COLOR);
  summary.setTabColor(VIEW_COLOR);
  wsSum.setTabColor(VIEW_COLOR);
  answers.setTabColor(DATA_COLOR);
  ws.setTabColor(DATA_COLOR);
  var old = ss.getSheetByName("満足度_旧");
  if (old) { old.setTabColor(DATA_COLOR); }

  // 手で編集しようとすると警告を出す（数式や記録をうっかり消さないため）
  [summary, wsSum, answers, ws].forEach(function (sh) {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
    sh.protect().setDescription("自動で記録・集計しています").setWarningOnly(true);
  });

  ss.setActiveSheet(summary);
}

function removeIfEmpty_(ss, name, headerRows) {
  var sh = ss.getSheetByName(name);
  if (sh && sh.getLastRow() <= headerRows && ss.getSheets().length > 1) { ss.deleteSheet(sh); }
}

// 記録用シートの見た目（見出し・日時の表示・列幅）
function styleDataSheet_(sh, widths, grayCols) {
  var cols = widths.length;
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, cols).setFontWeight("bold").setBackground(HEAD_BG).setWrap(true);
  sh.getRange("A2:A").setNumberFormat("yyyy/mm/dd hh:mm:ss");
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  (grayCols || []).forEach(function (c) {
    if (sh.getMaxRows() < 2) { return; }
    sh.getRange(2, c, sh.getMaxRows() - 1, 1).setFontColor("#9aa0a6");
  });
}

// ---------- 使い方 ----------
function setupHelp_(ss) {
  var sh = ss.getSheetByName("使い方") || ss.insertSheet("使い方");
  sh.clear();
  var rows = [
    ["多摩祭2026 動画コンテスト・ワークショップ アンケート 記録"],
    [""],
    ["■ 見るシート（青いタブ）"],
    ["集計　…　動画コンテストの順位。「面白かった」の人数が多い順。1位は黄色"],
    ["満足度集計　…　ワークショップ アンケートの人数・割合と、ひとことの一覧（右側）"],
    [""],
    ["■ 記録シート（灰色のタブ）"],
    ["回答　…　動画コンテストの回答が、サイトから自動で1件1行で入ります"],
    ["満足度　…　ワークショップ アンケートの回答が、サイトから自動で入ります（ひとことは H列）"],
    ["※ スタッフが「直前の1件を取り消す」を押した回答は、「取消」の印が付き、集計に数えません"],
    [""],
    ["■ 注意"],
    ["・どのシートも手で書き換えないでください（編集しようとすると警告が出ます）"],
    ["・集計シートの数式を消すと、順位や棒が出なくなります。崩れたときは Apps Script で setup を実行すると元に戻ります（記録は消えません）"],
    ["・テストで入れた行を消すときは、「回答」「満足度」の2行目以降を行ごと削除してください（1行目の見出しは残す）"],
    [""],
    ["■ リンク"],
    ["サイト：" + SITE_URL],
    ["動画フォルダ：https://drive.google.com/drive/folders/" + FOLDER_ID],
    ["（動画フォルダの直下に動画を入れると、サイトに自動で反映されます）"]
  ];
  sh.getRange(1, 1, rows.length, 1).setValues(rows);
  sh.getRange("A1").setFontSize(16).setFontWeight("bold");
  [3, 7, 12, 17].forEach(function (r) { sh.getRange(r, 1).setFontWeight("bold").setBackground(HEAD_BG); });
  sh.setColumnWidth(1, 900);
  sh.getRange(1, 1, rows.length, 1).setWrap(true).setVerticalAlignment("middle");
  return sh;
}

// ---------- 動画コンテスト ----------
function setupVideoAnswers_(ss) {
  var sh = ss.getSheetByName("回答") || ss.insertSheet("回答");
  if (sh.getLastRow() === 0) {
    sh.appendRow(["受信日時", "作品ID", "作品名", "面白かった(人)", "面白くなかった(人)", "端末送信日時", "送信ID"]);
  }
  styleDataSheet_(sh, [150, 120, 240, 110, 120, 190, 150], [2, 6, 7]);
  return sh;
}

function setupVideoSummary_(ss) {
  // 集計シート：「面白かった」の人数が多い順に並べ、順位と棒を付ける
  var sh = ss.getSheetByName("集計") || ss.insertSheet("集計");
  sh.clear();
  sh.clearConditionalFormatRules();

  // A 順位 / B〜F 作品ごとの合計（面白かった(人)が多い順、同数なら面白くなかった(人)が少ない順）
  sh.getRange("A1").setValue("順位");
  sh.getRange("B1").setFormula(
    '=IFERROR(QUERY(回答!A2:F, "select C, sum(D), sum(E), sum(D)+sum(E), count(C) where C is not null ' +
    "group by C order by sum(D) desc, sum(E) asc " +
    "label C '作品名', sum(D) '面白かった(人)', sum(E) '面白くなかった(人)', sum(D)+sum(E) '合計(人)', count(C) '上映回数'\", 0), " +
    '"まだ回答がありません")'
  );
  // 順位：面白かった(人)が多いほど上位。同数なら同じ順位
  sh.getRange("A2").setFormula('=ARRAYFORMULA(IF(C2:C="","",1+COUNTIF(C2:C,">"&C2:C)))');
  // G 面白かった率、H 棒（面白かった(人)の数だけ ■）
  sh.getRange("G1").setValue("面白かった率");
  sh.getRange("G2").setFormula('=ARRAYFORMULA(IF(C2:C="","",IFERROR(C2:C/E2:E,0)))');
  sh.getRange("H1").setValue("面白かった（棒）");
  sh.getRange("H2").setFormula('=ARRAYFORMULA(IF(C2:C="","",REPT("■",C2:C)))');

  // 全体の合計（J〜K列）
  sh.getRange("J1").setValue("全体の合計");
  sh.getRange("J2:J6").setValues([["総投票数(人)"], ["面白かった(人)"], ["面白くなかった(人)"], ["上映回数（回答の件数）"], ["投票のあった作品数"]]);
  sh.getRange("K2").setFormula('=SUM(E2:E)');
  sh.getRange("K3").setFormula('=SUM(C2:C)');
  sh.getRange("K4").setFormula('=SUM(D2:D)');
  sh.getRange("K5").setFormula('=SUM(F2:F)');
  sh.getRange("K6").setFormula('=COUNT(C2:C)');

  sh.getRange("A1:H1").setFontWeight("bold").setBackground(HEAD_BG).setWrap(true);
  sh.getRange("J1:K1").setFontWeight("bold").setBackground(HEAD_BG);
  sh.getRange("J2:J6").setFontWeight("bold");
  sh.getRange("K2:K6").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("right");
  sh.setFrozenRows(1);
  sh.getRange("A2:A").setHorizontalAlignment("center").setFontWeight("bold");
  sh.getRange("C2:C").setFontWeight("bold").setFontSize(14);
  sh.getRange("D2:E").setFontWeight("bold");
  sh.getRange("G2:G").setNumberFormat("0%");
  sh.getRange("H2:H").setFontColor("#2f6fed");
  [60, 240, 110, 120, 80, 80, 100, 320, 24, 190, 80].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  // 1位の行を目立たせる
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$A2=1")
      .setBackground("#fff2cc")
      .setBold(true)
      .setRanges([sh.getRange("A2:H300")])
      .build()
  ]);
  return sh;
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
var WS_HEADER = ["受信日時", "1 たのしかった？", "2 わかったこと（いくつでも）", "3 リサイクルしてみたい？", "端末送信日時", "送信ID", "取消", "4 ひとこと"];

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
  }
  if (sheet.getRange("H1").getValue() !== WS_HEADER[7]) { sheet.getRange("H1").setValue(WS_HEADER[7]); } // ひとこと の列（後から追加）
  styleDataSheet_(sheet, [150, 130, 420, 170, 190, 150, 60, 360], [5, 6]);
  sheet.getRange("H2:H").setWrap(true);
  sheet.getRange("C2:C").setWrap(true);
  sheet.getRange("G2:G").setFontColor("#d62828").setFontWeight("bold");

  // 満足度集計：質問ごとに、選択肢の人数・割合・棒と「有効回答／未回答／無効回答」
  var sum = ss.getSheetByName("満足度集計") || ss.insertSheet("満足度集計");
  sum.clear();
  sum.showColumns(1, sum.getMaxColumns());
  var G = '満足度!$G$2:$G';                 // 取消の列：「取消」が付いた回答は数えない
  var A = '満足度!$A$2:$A';                 // 受信日時（回答のある行）
  var cells = [];                            // [行, 列, 値または式]
  var titleRows = [], lineRows = [], noteRows = [], pctRows = [];
  var row = 1;
  var put = function (r, c, v) { cells.push([r, c, v]); };

  // F列（非表示）：割合のもとになる有効回答の件数
  var addGroup = function (title, col, items, multi) {
    var C = '満足度!$' + col + '$2:$' + col;
    var valid = 'COUNTIFS(' + C + ',"<>",' + C + ',"<>' + WS_INVALID + '",' + G + ',"<>取消")';
    var blank = 'COUNTIFS(' + A + ',"<>",' + C + ',"",' + G + ',"<>取消")';
    var bad = 'COUNTIFS(' + C + ',"' + WS_INVALID + '",' + G + ',"<>取消")';
    var head = row;
    put(row, 1, title); put(row, 2, "選択肢"); put(row, 3, "人数"); put(row, 4, "割合"); put(row, 5, "グラフ");
    put(row, 6, "=" + valid);
    titleRows.push(row); row++;
    items.forEach(function (item) {
      var crit = multi ? '"*"&"' + item + '"&"*"' : '"' + item + '"';
      put(row, 2, item);
      put(row, 3, '=COUNTIFS(' + C + ',' + crit + ',' + G + ',"<>取消")');
      put(row, 4, '=IFERROR(C' + row + '/$F$' + head + ',0)');
      put(row, 5, '=REPT("■",ROUND(D' + row + '*20))');
      pctRows.push(row); row++;
    });
    put(row, 2, '="有効回答："&' + valid + '&"件　未回答："&' + blank + '&"件　無効回答："&' + bad + '&"件"');
    lineRows.push(row); row++;
    if (multi) {
      put(row, 2, "※いくつでも えらべるため、割合の合計は100%をこえることがあります");
      noteRows.push(row); row++;
    }
    row++;
  };
  addGroup("1 ペパポンをつくって、たのしかった", "B", WS_Q1, false);
  addGroup("2 きょう、わかったことはどれかな（いくつでも）", "C", WS_Q2, true);
  addGroup("3 おうちのひとと、リサイクルをしてみたい", "D", WS_Q3, false);
  var total = row;
  put(total, 1, "回答の合計（人）");
  put(total, 3, '=COUNTIFS(' + A + ',"<>",' + G + ',"<>取消")');
  put(total + 1, 1, "取り消した回答（件）");
  put(total + 1, 3, '=COUNTIF(' + G + ',"取消")');

  cells.forEach(function (x) {
    var cell = sum.getRange(x[0], x[1]);
    if (typeof x[2] === "string" && x[2].charAt(0) === "=") { cell.setFormula(x[2]); } else { cell.setValue(x[2]); }
  });
  titleRows.forEach(function (r) {
    sum.getRange(r, 1, 1, 5).setFontWeight("bold").setBackground(HEAD_BG);
  });
  pctRows.forEach(function (r) {
    sum.getRange(r, 3).setFontWeight("bold").setFontSize(14);
    sum.getRange(r, 4).setNumberFormat("0%").setFontWeight("bold");
    sum.getRange(r, 5).setFontColor("#2f6fed");
  });
  lineRows.forEach(function (r) { sum.getRange(r, 2).setFontWeight("bold"); });
  noteRows.forEach(function (r) { sum.getRange(r, 2).setFontColor("#5f6368").setFontSize(9); });
  sum.getRange(total, 1, 2, 3).setFontWeight("bold");
  sum.getRange(total, 3).setFontSize(14);
  [330, 440, 70, 70, 220].forEach(function (w, i) { sum.setColumnWidth(i + 1, w); });
  sum.getRange(1, 3, total + 1, 2).setHorizontalAlignment("right");
  sum.hideColumns(6);

  // ひとこと（新しい順。取り消した回答は出さない）
  sum.getRange("H1").setValue("4 ひとこと（新しい順）").setFontWeight("bold").setBackground(HEAD_BG);
  sum.getRange("H2").setFormula(
    '=IFERROR(SORT(FILTER({満足度!A2:A, 満足度!H2:H}, 満足度!H2:H<>"", 満足度!G2:G<>"取消"), 1, FALSE), {"", "まだありません"})'
  );
  sum.getRange("H2:H").setNumberFormat("m/d hh:mm");
  sum.setColumnWidth(7, 24);
  sum.setColumnWidth(8, 110);
  sum.setColumnWidth(9, 420);
  sum.getRange("I2:I").setWrap(true);
  sum.getRange("H1:I1").merge();
  return sheet;
}

// テストで入れた行を消す（指定の送信ID と「check-」で始まる送信ID。何回実行してもよい）
function removeTestRows() {
  var ss = getSpreadsheet_();
  var ids = {
    "回答": ["6bdd4da8-676e-4fd8-baad-0cd2898032e0", "a269968b-3f08-426d-99bc-8645d8f90f3c", "500792ad-ccd9-4c57-9ddf-a41da017a43b"],
    "満足度": ["verify-1792671500407", "dualA1792671689253", "dualB1792671689253"]
  };
  var idCol = { "回答": 7, "満足度": 6 };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    Object.keys(ids).forEach(function (name) {
      var sh = ss.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) { return; }
      var vals = sh.getRange(2, idCol[name], sh.getLastRow() - 1, 1).getValues();
      for (var i = vals.length - 1; i >= 0; i--) {          // 下の行から消す
        var v = String(vals[i][0]);
        // 指定の送信ID、または「check-」で始まる送信ID（動作確認用）を消す
        if (ids[name].indexOf(v) >= 0 || v.indexOf("check-") === 0) { sh.deleteRow(i + 2); }
      }
    });
  } finally {
    lock.releaseLock();
  }
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
          var cm = String(r.comment || "").replace(/^[=+\-@]/, "'$&").slice(0, 500); // 数式として扱われないように
          if (a1 === "" && a2 === "" && a3 === "" && cm === "") { return; }
          workshop.push({ id: wid, q1: a1, q2: a2, q3: a3, comment: cm, sentAt: wsent });
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
          wrows.push([now, r.q1, r.q2, r.q3, r.sentAt, r.id, "", r.comment]);
        });
        if (wrows.length > 0) {
          ws.getRange(wlast + 1, 1, wrows.length, 8).setValues(wrows);
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
