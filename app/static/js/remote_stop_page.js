/*
  remote_stop_page.js — 管理者以外停止のお願いのあいだ、この画面の「打ちかけ」を帯(remote_stop_band.js)に渡す

  UF検査判定は操作のたびに手元DBへ書くので、消えるのは**画面にしかない入力**だけ:
  欄に打ったが、まだ Enter・確定(change)していないもの。

  - window.__remoteStopUnsaved() … 打ちかけの欄の呼び名(帯が問い合わせのたびに言い添える → 利用者を待つ)
  - window.__remoteStopFlush()   … 止めると決まったとき、打ちかけを値ごとサーバへ送る(手元の控えに書かれ、
                                   次の起動で知らせる)。そのあと打っている欄を離れる
  パスワードの欄は数えない・送らない。
*/
(function () {
  "use strict";
  var dirty = new Set();
  var PAGE = location.pathname;

  function isField(t) {
    if (!t || !t.tagName) return false;
    if (t.readOnly || t.disabled) return false;
    if (t.tagName === "TEXTAREA") return true;
    if (t.tagName !== "INPUT") return false;
    return ["", "text", "number", "search", "tel", "email", "url"].indexOf((t.getAttribute("type") || "").toLowerCase()) >= 0;
  }

  function label(t) {
    var text = t.getAttribute("aria-label") || "";
    if (!text && t.labels && t.labels.length) text = t.labels[0].textContent || "";
    if (!text) text = t.getAttribute("placeholder") || t.name || t.id || "入力欄";
    return text.replace(/\s+/g, " ").trim().slice(0, 60);
  }

  function live() {
    var out = [];
    dirty.forEach(function (t) {
      if (!t.isConnected || !String(t.value || "").trim()) { dirty.delete(t); return; }
      out.push(t);
    });
    return out;
  }

  document.addEventListener("input", function (e) { if (isField(e.target)) dirty.add(e.target); }, true);
  document.addEventListener("change", function (e) { dirty.delete(e.target); }, true);
  document.addEventListener("keydown", function (e) { if (e.key === "Enter") dirty.delete(e.target); }, true);

  window.__remoteStopUnsaved = function () {
    return live().map(function (t) { return "入力しかけの欄「" + label(t) + "」"; });
  };

  window.__remoteStopFlush = function () {
    var fields = live();
    var drafts = fields.map(function (t) { return { label: label(t), value: String(t.value || "") }; });
    var active = document.activeElement;
    if (active && active !== document.body && typeof active.blur === "function") active.blur();
    if (!drafts.length) return Promise.resolve();
    var token = (window.APP && window.APP.token) || "";
    return fetch("/api/remote_stop/drafts", {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Tool-Token": token },
      body: JSON.stringify({ page: PAGE, drafts: drafts })
    }).then(function () { fields.forEach(function (t) { dirty.delete(t); }); }).catch(function () {});
  };
})();
