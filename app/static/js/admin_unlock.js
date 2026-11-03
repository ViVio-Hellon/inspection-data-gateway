/*
  admin_unlock.js — マスタ編集の鍵(管理者パスワード)。2026-11-02 の決め

  マスタ管理・保守の画面は誰でも開けるが、編集は「管理者パスワード」と「アクセス権限の管理者」の両方が要る。
  判断はサーバ(app/routes/admin_auth.py → ufkensa.auth)。ここは窓を出して送るだけ。

  使い方:
      import { ensureUnlocked, withUnlock, mountLockBadge } from "../admin_unlock.js";
      if (!(await ensureUnlocked())) return;                 // 鍵が閉じていればパスワードの窓
      const d = await withUnlock(() => api.post(url, body)); // 403 {code:"locked"} なら窓 → 呼び直す(1 回)
      mountLockBadge(document.getElementById("lockBox"));   // 「鍵: 閉 / 開」と 開ける・閉じる

  鍵の状態が変わると window に "admin-lock-change"(detail = 状態)を流す。
  ブラウザの prompt は使わない(askbox.js の注のとおり。デスクトップ版で窓が出ないことがある)。
*/

import { api } from "./api.js";
import { toast } from "./toast.js";

export const EVENT = "admin-lock-change";
let last = null;

function emit(state) {
  if (!state) return;
  last = state;
  window.dispatchEvent(new CustomEvent(EVENT, { detail: state }));
}

/** いまの鍵の状態 {unlocked, can_edit, admin, reason, custom, min_length}。 */
export async function adminState() {
  const s = await api.get("/api/admin/state");
  emit(s);
  return s;
}

export function lastState() { return last; }

/* ---------- パスワードの窓(type=password。間違えても閉じずにその場で言う) ---------- */
let dlg = null;

function ensureDialog() {
  if (dlg && document.body.contains(dlg.dialog)) return dlg;
  const style = document.createElement("style");
  style.textContent = `
    .unlockbox { border: 1px solid var(--rule); border-radius: var(--radius, 10px); padding: 0;
                 background: var(--surface); color: var(--ink); width: min(440px, 92vw);
                 box-shadow: var(--shadow-pop, 0 10px 30px rgb(0 0 0 / 30%)); }
    .unlockbox::backdrop { background: rgb(0 0 0 / 45%); }
    .unlockbox__body { padding: 18px 20px 6px; display: flex; flex-direction: column; gap: 10px; }
    .unlockbox__title { margin: 0; font-size: var(--fs-md, 14px); font-weight: 700; }
    .unlockbox__text { margin: 0; font-size: var(--fs-sm, 12px); color: var(--muted); line-height: 1.6; white-space: pre-wrap; }
    .unlockbox__err { margin: 0; min-height: 1.4em; font-size: var(--fs-sm, 12px); color: var(--state-error); font-weight: 700; }
    .unlockbox__foot { display: flex; justify-content: flex-end; gap: 8px; padding: 8px 20px 16px; }
  `;
  const dialog = document.createElement("dialog");
  dialog.className = "unlockbox";
  dialog.innerHTML = `
    <form method="dialog" class="unlockbox__form" autocomplete="off">
      <div class="unlockbox__body">
        <p class="unlockbox__title">🔒 編集の鍵を開ける</p>
        <p class="unlockbox__text"></p>
        <input class="input" type="password" name="password" autocomplete="current-password" placeholder="管理者パスワード">
        <p class="unlockbox__err" role="alert"></p>
      </div>
      <div class="unlockbox__foot">
        <button class="btn" type="button" data-answer="no">やめる</button>
        <button class="btn btn--commit" type="submit" data-answer="yes">開ける</button>
      </div>
    </form>`;
  document.head.appendChild(style);
  document.body.appendChild(dialog);
  dlg = {
    dialog,
    form: dialog.querySelector("form"),
    text: dialog.querySelector(".unlockbox__text"),
    input: dialog.querySelector("input"),
    err: dialog.querySelector(".unlockbox__err"),
    no: dialog.querySelector('[data-answer="no"]'),
    yes: dialog.querySelector('[data-answer="yes"]'),
  };
  return dlg;
}

function askPassword(message) {
  const b = ensureDialog();
  if (b.dialog.open) b.dialog.close();
  b.text.textContent = message;
  b.input.value = "";
  b.err.textContent = "";
  return new Promise((resolve) => {
    let busy = false;
    const finish = (ok) => {
      b.form.removeEventListener("submit", onSubmit);
      b.no.removeEventListener("click", onNo);
      b.dialog.removeEventListener("close", onClose);
      b.dialog.removeEventListener("click", onBackdrop);
      if (b.dialog.open) b.dialog.close();
      b.input.value = "";
      resolve(ok);
    };
    const onSubmit = async (e) => {
      e.preventDefault();
      if (busy) return;
      busy = true;
      b.yes.disabled = true;
      try {
        const d = await api.post("/api/admin/unlock", { password: b.input.value });
        emit(d.state);
        if (d.ok) { finish(true); return; }
        b.err.textContent = d.reason || "開けられませんでした";
        b.input.select();
        if (d.state && !d.state.admin) { b.input.disabled = true; }
      } catch (err) {
        b.err.textContent = err.message || String(err);
      } finally {
        busy = false;
        b.yes.disabled = false;
      }
    };
    const onNo = () => finish(false);
    const onClose = () => finish(false);                     // Esc
    const onBackdrop = (e) => { if (e.target === b.dialog) finish(false); };
    b.form.addEventListener("submit", onSubmit);
    b.no.addEventListener("click", onNo);
    b.dialog.addEventListener("close", onClose);
    b.dialog.addEventListener("click", onBackdrop);
    b.input.disabled = false;
    b.dialog.showModal();
    b.input.focus();
  });
}

/**
 * 編集してよい状態にする。もう開いていれば true。
 * 管理者の権限が無ければ窓は出さずに理由を知らせて false。やめたら false。
 */
export async function ensureUnlocked(message = "") {
  let s;
  try { s = await adminState(); } catch (err) { toast(err.message || String(err), "error"); return false; }
  if (s.can_edit) return true;
  if (!s.admin) { toast(s.reason || "編集には、アクセス権限の「管理者」が要ります", "error"); return false; }
  const ok = await askPassword(message || (s.reason || "編集するには、管理者パスワードを入れてください")
                               + "\n開けると、このアプリを閉じるか「閉じる」を押すまで編集できます。");
  if (ok) toast("編集の鍵を開けました", "ok");
  return ok;
}

/** fn() が 403 {code:"locked"} で断られたら、鍵の窓を出して 1 回だけ呼び直す。やめたら null。 */
export async function withUnlock(fn) {
  try {
    return await fn();
  } catch (err) {
    if (err && err.status === 403 && err.code === "locked") {
      if (!(await ensureUnlocked(err.message))) return null;
      return await fn();
    }
    throw err;
  }
}

/** 鍵を閉じる。 */
export async function lockNow() {
  const d = await api.post("/api/admin/lock", {});
  emit(d.state);
  toast("編集の鍵を閉じました", "ok");
  return d.state;
}

/**
 * 「鍵: 閉 / 開」の札と 開ける・閉じる ボタン。container の中身を置き換える。
 * 状態が変われば(どこで開けても)描き直す。
 */
export function mountLockBadge(container) {
  if (!container) return;
  container.classList.add("lockbadge");
  const chip = document.createElement("span");
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "btn btn--sm";
  const why = document.createElement("span");
  why.className = "lockbadge__why muted";
  container.replaceChildren(chip, btn, why);
  const render = (s) => {
    if (!s) return;
    chip.className = "lv " + (s.can_edit ? "lv--ok" : "lv--warn");
    chip.textContent = s.can_edit ? "🔓 鍵: 開" : "🔒 鍵: 閉";
    btn.textContent = s.can_edit ? "閉じる" : "開ける";
    btn.hidden = !s.admin && !s.unlocked;
    why.textContent = s.can_edit ? "編集できます" : (s.reason || "");
  };
  btn.addEventListener("click", async () => {
    try {
      if (last && last.can_edit) await lockNow();
      else await ensureUnlocked();
    } catch (err) { toast(err.message || String(err), "error"); }
  });
  window.addEventListener(EVENT, (e) => render(e.detail));
  if (last) render(last);
  adminState().catch(() => {});
}
