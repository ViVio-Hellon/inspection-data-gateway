/*
  screen.js — 操作できる画面は1つだけ(2026-11-11)

  現場の声「まだ多重起動できる」。起動をもう一度押すとブラウザの画面がもう1つ開き、
  どちらでも操作できた。そこで:

  - 数秒ごとに `GET /api/screen` を送る(心拍)。サーバは「いま使っている画面か」と
    「前に出て」の番号(raise_seq)を返す
  - 2回目の起動は、新しく開く代わりに「前に出て」を出す。受け取ったら window.focus()・
    題の点滅・知らせ「すでに開いています(もう一度起動されました)」。
    **ブラウザは、操作していない窓を前に出すこと(focus)をたいてい許さない。**
    そのため点滅と知らせで気づけるようにしている(docs/06_意図した差異.md 2026-11-11)
  - 同じブラウザのほかのタブとは BroadcastChannel(使えなければ localStorage)で
    すぐに知らせ合う。サーバの判定が正(別のブラウザでも効く)
  - 使っていない画面には覆い「この画面はほかのウィンドウで開いています」を出す。
    「こちらで使う」で取り返す(読み込み直して最新の状態で出す)/「閉じる」
*/

import { api } from "./api.js";
import { toast } from "./toast.js";

const HANDOFF_KEY = "ufkensa.screen.handoff";
const CHANNEL = "ufkensa.screen";
const STORAGE_KEY = "ufkensa.screen.claim";

let lastRaise = null;
let overlay = null;
let channel = null;
let flashTimer = null;

function me() { return window.APP.screenId || ""; }

function visible() { return document.visibilityState === "visible"; }

/* ---- 覆い ------------------------------------------------------ */
function buildOverlay() {
  const el = document.createElement("div");
  el.className = "taken";
  el.id = "screenLock";
  el.setAttribute("role", "alertdialog");
  el.setAttribute("aria-modal", "true");
  el.setAttribute("aria-labelledby", "screenLockText");
  el.hidden = true;
  el.innerHTML = `
    <div class="taken__card">
      <strong class="taken__title" id="screenLockText">この画面はほかのウィンドウで開いています</strong>
      <p class="taken__note">操作できる画面は1つだけです。この画面で続けるときは「こちらで使う」を押してください(ほかの画面は操作できなくなります)。</p>
      <p class="taken__why" data-screen="note" hidden>このタブはブラウザの決まりで閉じられませんでした。タブの × で閉じてください。</p>
      <div class="taken__acts">
        <button class="btn btn--commit" type="button" data-screen="take">こちらで使う</button>
        <button class="btn" type="button" data-screen="close">閉じる</button>
      </div>
    </div>`;
  el.querySelector('[data-screen="take"]').addEventListener("click", takeOver);
  el.querySelector('[data-screen="close"]').addEventListener("click", () => {
    try { window.close(); } catch { /* 閉じられない */ }
    // スクリプトで開いたタブでなければ閉じられない。少し待って残っていれば案内する
    setTimeout(() => { el.querySelector('[data-screen="note"]').hidden = false; }, 300);
  });
  document.body.appendChild(el);
  return el;
}

function setInactive(inactive) {
  if (!overlay) overlay = buildOverlay();
  if (overlay.hidden === !inactive) return;
  overlay.hidden = !inactive;
  document.documentElement.classList.toggle("screen-inactive", inactive);
  if (inactive) overlay.querySelector('[data-screen="take"]').focus();
}

export function isInactive() { return !!overlay && !overlay.hidden; }

/* ---- 取り返す --------------------------------------------------- */
async function takeOver() {
  try {
    await api.post("/api/screen/claim", {});
  } catch (e) {
    toast(e.message || "切り替えられませんでした", "error");
    return;
  }
  announce();
  // ほかの画面で進んだ状態を出すため、読み込み直す(番号は引き継ぐ)
  location.reload();
}

/* ---- 同じブラウザのほかのタブへ ----------------------------------- */
function announce() {
  const msg = { type: "claimed", id: me(), at: Date.now() };
  try { if (channel) channel.postMessage(msg); } catch { /* 送れない */ }
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(msg)); } catch { /* 使えない */ }
}

function heard(msg) {
  if (msg && msg.type === "claimed" && msg.id && msg.id !== me()) setInactive(true);
}

function listen() {
  try {
    if ("BroadcastChannel" in window) {
      channel = new BroadcastChannel(CHANNEL);
      channel.onmessage = (ev) => heard(ev.data);
    }
  } catch { channel = null; }
  try {
    window.addEventListener("storage", (ev) => {
      if (ev.key !== STORAGE_KEY || !ev.newValue) return;
      try { heard(JSON.parse(ev.newValue)); } catch { /* 壊れている */ }
    });
  } catch { /* 使えない */ }
}

/* ---- 前に出る ---------------------------------------------------- */
function bringToFront() {
  try { window.focus(); } catch { /* 許されない */ }
  toast("すでに開いています(もう一度起動されました)", "info");
  if (flashTimer) return;
  const title = document.title;
  let n = 0;
  flashTimer = setInterval(() => {
    document.title = n % 2 ? title : "★ ここで開いています";
    if (++n >= 12 || (document.hasFocus() && n >= 4)) {
      clearInterval(flashTimer);
      flashTimer = null;
      document.title = title;
    }
  }, 700);
}

/* ---- 心拍 ------------------------------------------------------- */
async function ping() {
  let res;
  try {
    res = await api.get(`/api/screen?visible=${visible() ? 1 : 0}`);
  } catch { return; }
  if (!res) return;
  // 覆いを出していた画面が、ほかの画面が閉じたので使えるようになった。覆いの間に断られた
  // 読み込み(開いた直後の操作など)があるので、読み込み直して最新の状態で出す
  if (isInactive() && res.active !== false) { location.reload(); return; }
  setInactive(res.active === false);
  if (typeof res.raise_seq === "number") {
    if (lastRaise !== null && res.raise_seq > lastRaise && res.active !== false) bringToFront();
    lastRaise = res.raise_seq;
  }
}

/* 断られた(ほかの画面が使っている)。api.js が知らせてくる */
function onRefused() { setInactive(true); }

export function start() {
  if (!me()) return;
  listen();
  document.addEventListener("app:inactive-screen", onRefused);
  // 同じタブで読み込み直すときは番号を引き継がせる(base.html が次のページで読む)
  window.addEventListener("pagehide", () => {
    try { sessionStorage.setItem(HANDOFF_KEY, me()); } catch { /* 使えない */ }
  });
  document.addEventListener("visibilitychange", () => { if (visible()) ping(); });
  ping();
  setInterval(ping, window.APP.screenPollMs || 3000);
}
