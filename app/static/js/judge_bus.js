/*
  判定を別窓へ渡す(2026-11-26 現場の声「キャスク・全ヵ所の判定は別窓で大きく、VBA のように何枚目・判定が何か」)。
  同じアプリの窓どうしで渡すだけ(サーバへは聞かない)。BroadcastChannel、だめなら localStorage の知らせ。
  info = { screen: "cask" | "zenkasho", no: "12", label: "厚2", mark: "○" | "×", at: 時刻 }
*/
const NAME = "ufkensa-judge";
const KEY = "ufkensa.judge.last";
let ch = null;
try { ch = new BroadcastChannel(NAME); } catch { ch = null; }

export function publishJudge(info) {
  const msg = { ...info, at: Date.now() };
  try { if (ch) ch.postMessage(msg); } catch { /* 渡せなくても検査は止めない */ }
  try { localStorage.setItem(KEY, JSON.stringify(msg)); } catch { /* 同上 */ }
}

export function lastJudge() {
  try { return JSON.parse(localStorage.getItem(KEY) || "null"); } catch { return null; }
}

export function onJudge(fn) {
  let seen = 0;
  const take = (m) => { if (m && m.at && m.at !== seen) { seen = m.at; fn(m); } };
  if (ch) ch.addEventListener("message", (e) => take(e.data));
  window.addEventListener("storage", (e) => { if (e.key === KEY) { try { take(JSON.parse(e.newValue)); } catch { /* 壊れた知らせ */ } } });
}
