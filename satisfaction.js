(function () {
  "use strict";

  var CFG = window.APP_CONFIG || {};
  var ENDPOINT = CFG.SURVEY_ENDPOINT || "";
  var $ = function (id) { return document.getElementById(id); };

  var answers = { q1: "", q2: "", q3: "" };
  var lastId = null;   // 直前に送った1件（取り消し用）

  var QUEUE_KEY = "tamasai_ws_queue";
  var COUNT_KEY = "tamasai_ws_count";
  var queueMem = [];
  var flushing = false;
  var retryTimer = null;
  var retryDelay = 3000;
  var toastTimer = null;

  // ---------- 画面 ----------
  function newId() {
    if (window.crypto && window.crypto.randomUUID) { return window.crypto.randomUUID(); }
    return "id" + Date.now() + Math.random().toString(16).slice(2);
  }

  function getCount() {
    try { return parseInt(localStorage.getItem(COUNT_KEY) || "0", 10) || 0; } catch (e) { return 0; }
  }
  function setCount(n) {
    try { localStorage.setItem(COUNT_KEY, String(Math.max(0, n))); } catch (e) { /* 保存できなくても続行 */ }
    $("ws-count").textContent = "この端末の記録： " + Math.max(0, n) + "件";
  }

  function toast(msg) {
    var el = $("ws-toast");
    el.textContent = msg;
    el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.hidden = true; }, 4000);
  }

  function updateSummary() {
    var parts = [answers.q1, answers.q2, answers.q3].filter(function (v) { return v; });
    var comment = $("ws-comment").value.trim();
    $("ws-summary").textContent = parts.length === 0
      ? ""
      : "えらんだ内容： " + [answers.q1 || "①まだ", answers.q2 || "②まだ", answers.q3 || "③まだ"].join(" ／ ") +
        (comment ? " ／ ひとこと有り" : "");
  }

  function resetForm() {
    answers = { q1: "", q2: "", q3: "" };
    var btns = document.querySelectorAll(".choice");
    for (var i = 0; i < btns.length; i++) { btns[i].setAttribute("aria-pressed", "false"); }
    $("ws-comment").value = "";
    $("ws-error").hidden = true;
    updateSummary();
    window.scrollTo(0, 0);
  }

  function onChoice(ev) {
    var btn = ev.target.closest ? ev.target.closest(".choice") : null;
    if (!btn) { return; }
    var group = btn.parentNode;
    var q = group.getAttribute("data-q");
    answers[q] = btn.getAttribute("data-v");
    var sibs = group.querySelectorAll(".choice");
    for (var i = 0; i < sibs.length; i++) {
      sibs[i].setAttribute("aria-pressed", sibs[i] === btn ? "true" : "false");
    }
    $("ws-error").hidden = true;
    updateSummary();
  }

  // ---------- 送信（動画コンテストと同じ：裏で送り、失敗したら自動で再送） ----------
  function persistQueue() {
    try { localStorage.setItem(QUEUE_KEY, JSON.stringify(queueMem)); } catch (e) { /* 続行 */ }
  }
  function loadQueue() {
    try { return JSON.parse(localStorage.getItem(QUEUE_KEY) || "[]"); } catch (e) { return []; }
  }

  function updateQueueStatus() {
    var n = queueMem.length;
    var bar = $("queue-status");
    bar.hidden = n === 0;
    bar.textContent = "送信待ち " + n + "件（自動で再送します）";
  }

  function scheduleRetry() {
    if (retryTimer) { return; }
    retryTimer = setTimeout(function () { retryTimer = null; flush(); }, retryDelay);
    retryDelay = Math.min(retryDelay * 2, 30000);
  }

  function flush() {
    if (flushing || !ENDPOINT || queueMem.length === 0) { return; }
    flushing = true;
    var batch = queueMem.slice(0, 20);
    var ctrl = new AbortController();
    var timer = setTimeout(function () { ctrl.abort(); }, 25000);

    fetch(ENDPOINT, {
      method: "POST",
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

  function enqueue(rec) {
    if (!ENDPOINT) {
      // 記録先が未設定：この端末にだけ保存（動作確認用）
      try {
        var key = "tamasai_ws_local";
        var list = JSON.parse(localStorage.getItem(key) || "[]");
        list.push(rec);
        localStorage.setItem(key, JSON.stringify(list));
      } catch (e) { /* 続行 */ }
      return;
    }
    queueMem.push(rec);
    persistQueue();
    updateQueueStatus();
    flush();
  }

  function submit(ev) {
    ev.preventDefault();
    var err = $("ws-error");
    if (!answers.q1 || !answers.q2 || !answers.q3) {
      err.textContent = "①②③を えらんでください。";
      err.hidden = false;
      return;
    }
    var id = newId();
    enqueue({
      kind: "workshop",
      submissionId: id,
      q1: answers.q1,
      q2: answers.q2,
      q3: answers.q3,
      comment: $("ws-comment").value.trim(),
      sentAt: new Date().toISOString()
    });
    lastId = id;
    $("ws-undo").disabled = false;
    setCount(getCount() + 1);
    resetForm();
    toast("きろくしました。つぎの人 どうぞ！");
  }

  function undo() {
    if (!lastId) { return; }
    var id = lastId;
    lastId = null;
    $("ws-undo").disabled = true;

    // まだ送っていない（送信中でもない）なら、送信待ちから取り除くだけでよい
    var idx = -1;
    for (var i = 0; i < queueMem.length; i++) {
      if (queueMem[i].submissionId === id) { idx = i; break; }
    }
    if (idx >= 0 && !flushing) {
      queueMem.splice(idx, 1);
      persistQueue();
      updateQueueStatus();
    } else {
      // すでに送った、または送信中：取り消しの連絡を送る
      enqueue({ kind: "workshop", op: "cancel", submissionId: newId(), targetId: id, sentAt: new Date().toISOString() });
    }
    setCount(getCount() - 1);
    toast("直前の1件を取り消しました。");
  }

  function init() {
    var groups = document.querySelectorAll(".choices");
    for (var i = 0; i < groups.length; i++) { groups[i].addEventListener("click", onChoice); }
    $("ws-comment").addEventListener("input", updateSummary);
    $("ws-form").addEventListener("submit", submit);
    $("ws-undo").addEventListener("click", undo);

    setCount(getCount());
    queueMem = loadQueue();
    updateQueueStatus();
    flush();
    window.addEventListener("online", flush);
    document.addEventListener("visibilitychange", function () { if (!document.hidden) { flush(); } });

    if (!ENDPOINT) {
      var n = $("notice");
      n.textContent = "記録先が未設定です。回答はこの端末にだけ保存されます（動作確認用）。";
      n.hidden = false;
    }
  }

  init();
})();
