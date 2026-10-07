(function () {
  "use strict";

  var CFG = window.APP_CONFIG || {};
  var MAX_PEOPLE = CFG.MAX_PEOPLE || 10;
  var ENDPOINT = CFG.SURVEY_ENDPOINT || "";

  var $ = function (id) { return document.getElementById(id); };
  var screens = {
    start: $("screen-start"),
    play: $("screen-play"),
    survey: $("screen-survey"),
    end: $("screen-end")
  };

  var works = [];     // works.json の全作品
  var queue = [];     // シャッフルした再生順
  var pos = 0;        // 今の再生位置（0始まり）
  var sending = false;

  function show(name) {
    Object.keys(screens).forEach(function (k) { screens[k].hidden = (k !== name); });
    window.scrollTo(0, 0);
  }

  function notice(msg) {
    var el = $("notice");
    el.textContent = msg || "";
    el.hidden = !msg;
  }

  // Fisher–Yates
  function shuffle(arr) {
    var a = arr.slice();
    for (var i = a.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var t = a[i]; a[i] = a[j]; a[j] = t;
    }
    return a;
  }

  function fillSelect(sel) {
    sel.innerHTML = "";
    for (var n = 0; n <= MAX_PEOPLE; n++) {
      var o = document.createElement("option");
      o.value = String(n);
      o.textContent = n + "人";
      sel.appendChild(o);
    }
  }

  function progressText() {
    return (pos + 1) + " / " + queue.length + "本目";
  }

  function renderPlayer(work) {
    var box = $("player");
    box.innerHTML = "";
    // 向き：works.json の "orientation"（"portrait"=縦 / "landscape"=横）。省略時は縦
    box.className = "player " + (work.orientation === "landscape" ? "landscape" : "portrait");
    if (/^DUMMY/.test(work.driveId)) {
      // 動作確認用のダミー表示（本番では本物の driveId に差し替える）
      var ph = document.createElement("div");
      ph.className = "placeholder";
      var s = document.createElement("strong");
      s.textContent = work.title;
      var p = document.createElement("span");
      p.textContent = "（ダミー動画：works.json の driveId を本物に差し替えてください）";
      ph.appendChild(s);
      ph.appendChild(p);
      box.appendChild(ph);
      return;
    }
    var f = document.createElement("iframe");
    f.src = "https://drive.google.com/file/d/" + encodeURIComponent(work.driveId) + "/preview";
    f.title = work.title;
    f.allow = "autoplay; fullscreen";
    f.allowFullscreen = true;
    box.appendChild(f);
  }

  function playCurrent() {
    var work = queue[pos];
    $("play-progress").textContent = progressText();
    renderPlayer(work);
    show("play");
  }

  function startRound() {
    queue = shuffle(works);
    pos = 0;
    playCurrent();
  }

  function goSurvey() {
    $("player").innerHTML = ""; // 再生を止める
    $("survey-progress").textContent = progressText();
    $("sel-funny").value = "0";
    $("sel-notfunny").value = "0";
    $("survey-error").hidden = true;
    $("btn-submit").disabled = false;
    $("btn-submit").textContent = "送信";
    show("survey");
  }

  function afterAnswer() {
    pos += 1;
    if (pos >= queue.length) {
      show("end");
    } else {
      playCurrent();
    }
  }

  function saveLocal(rec) {
    try {
      var key = "tamasai_survey_local";
      var list = JSON.parse(localStorage.getItem(key) || "[]");
      list.push(rec);
      localStorage.setItem(key, JSON.stringify(list));
    } catch (e) { /* 保存できなくても続行 */ }
  }

  function submitSurvey(ev) {
    ev.preventDefault();
    if (sending) { return; }
    var funny = parseInt($("sel-funny").value, 10);
    var notFunny = parseInt($("sel-notfunny").value, 10);
    var err = $("survey-error");

    if (funny + notFunny < 1) {
      err.textContent = "人数を1人以上選んでください。";
      err.hidden = false;
      return;
    }
    err.hidden = true;

    var work = queue[pos];
    var rec = {
      workId: work.id,
      title: work.title,
      funny: funny,
      notFunny: notFunny,
      sentAt: new Date().toISOString()
    };

    if (!ENDPOINT) {
      // 記録先が未設定：この端末にだけ保存して先へ進む（動作確認用）
      saveLocal(rec);
      afterAnswer();
      return;
    }

    sending = true;
    var btn = $("btn-submit");
    btn.disabled = true;
    btn.textContent = "送信中…";

    fetch(ENDPOINT, {
      method: "POST",
      // text/plain にすると CORS のプリフライトが発生しない
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(rec)
    })
      .then(function (res) {
        if (!res.ok) { throw new Error("HTTP " + res.status); }
        return res.json();
      })
      .then(function (json) {
        if (!json || json.ok !== true) { throw new Error((json && json.error) || "unknown"); }
        sending = false;
        afterAnswer();
      })
      .catch(function () {
        sending = false;
        btn.disabled = false;
        btn.textContent = "もう一度送信";
        err.textContent = "送信できませんでした。通信状況を確認して、もう一度押してください。";
        err.hidden = false;
      });
  }

  function init() {
    fillSelect($("sel-funny"));
    fillSelect($("sel-notfunny"));

    $("btn-start").addEventListener("click", startRound);
    $("btn-to-survey").addEventListener("click", goSurvey);
    $("survey-form").addEventListener("submit", submitSurvey);
    $("btn-restart").addEventListener("click", function () { show("start"); });

    if (!ENDPOINT) {
      notice("記録先が未設定です。回答はこの端末にだけ保存されます（動作確認用）。");
    }

    fetch("works.json?t=" + Date.now(), { cache: "no-store" })
      .then(function (r) {
        if (!r.ok) { throw new Error("HTTP " + r.status); }
        return r.json();
      })
      .then(function (list) {
        works = (Array.isArray(list) ? list : []).filter(function (w) {
          return w && w.id && w.title && w.driveId;
        });
        if (works.length === 0) {
          notice("works.json に動画が登録されていません。");
          return;
        }
        $("start-count").textContent = "全" + works.length + "本";
        show("start");
      })
      .catch(function () {
        notice("works.json を読み込めませんでした。");
      });
  }

  init();
})();
