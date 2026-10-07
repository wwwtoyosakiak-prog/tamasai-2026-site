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

  // ---------- アンケートの送信（裏で送って、失敗したら自動で再送する） ----------
  var QUEUE_KEY = "tamasai_survey_queue";
  var queueMem = [];
  var flushing = false;
  var retryTimer = null;
  var retryDelay = 3000;

  function persistQueue() {
    try { localStorage.setItem(QUEUE_KEY, JSON.stringify(queueMem)); } catch (e) { /* 保存できなくても続行 */ }
  }

  function loadQueue() {
    try { return JSON.parse(localStorage.getItem(QUEUE_KEY) || "[]"); } catch (e) { return []; }
  }

  function newId() {
    if (window.crypto && window.crypto.randomUUID) { return window.crypto.randomUUID(); }
    return "id" + Date.now() + Math.random().toString(16).slice(2);
  }

  function updateQueueStatus() {
    var n = queueMem.length;
    var bar = $("queue-status");
    bar.hidden = n === 0;
    bar.textContent = "送信待ち " + n + "件（自動で再送します）";
    var end = $("end-pending");
    end.hidden = n === 0;
    end.textContent = "未送信の回答が " + n + "件あります。このページを開いたまま、通信できる場所で少しお待ちください。";
  }

  function scheduleRetry() {
    if (retryTimer) { return; }
    retryTimer = setTimeout(function () { retryTimer = null; flush(); }, retryDelay);
    retryDelay = Math.min(retryDelay * 2, 30000);
  }

  function flush() {
    if (flushing || !ENDPOINT || queueMem.length === 0) { return; }
    flushing = true;
    // たまっている回答を、まとめて1回で送る（1回あたりの待ち時間が長いため）
    var batch = queueMem.slice(0, 20);
    var ctrl = new AbortController();
    var timer = setTimeout(function () { ctrl.abort(); }, 25000);

    fetch(ENDPOINT, {
      method: "POST",
      // text/plain にすると CORS のプリフライトが発生しない
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ records: batch }),
      signal: ctrl.signal
    })
      .then(function (res) {
        if (!res.ok) { throw new Error("HTTP " + res.status); }
        return res.json();
      })
      .then(function (json) {
        clearTimeout(timer);
        // 成功、または二度と通らない回答（人数が不正）は、キューから外す
        if (json && (json.ok === true || json.error === "invalid count")) {
          queueMem.splice(0, batch.length);
          persistQueue();
          updateQueueStatus();
          retryDelay = 3000;
          flushing = false;
          flush();
        } else {
          throw new Error((json && json.error) || "unknown");
        }
      })
      .catch(function () {
        clearTimeout(timer);
        flushing = false;
        scheduleRetry();
      });
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
      submissionId: newId(),
      workId: work.id,
      title: work.title,
      funny: funny,
      notFunny: notFunny,
      sentAt: new Date().toISOString()
    };

    if (!ENDPOINT) {
      // 記録先が未設定：この端末にだけ保存して先へ進む（動作確認用）
      saveLocal(rec);
    } else {
      queueMem.push(rec);
      persistQueue();
      updateQueueStatus();
      flush();
    }
    // 返事を待たずに、すぐ次へ進む
    afterAnswer();
  }

  function init() {
    fillSelect($("sel-funny"));
    fillSelect($("sel-notfunny"));

    $("btn-start").addEventListener("click", startRound);
    $("btn-to-survey").addEventListener("click", goSurvey);
    $("survey-form").addEventListener("submit", submitSurvey);
    $("btn-restart").addEventListener("click", function () { show("start"); });

    // 前回送れなかった回答があれば、続きを送る
    queueMem = loadQueue();
    updateQueueStatus();
    flush();
    window.addEventListener("online", flush);
    document.addEventListener("visibilitychange", function () { if (!document.hidden) { flush(); } });

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
