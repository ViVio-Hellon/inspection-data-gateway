/*
  読み上げ(キャスクの UFCask.Read と CommandButton14 の代わり。2026-11-08)。

  VBA は Excel の Range.Speak で読んでいた(読み終わるまで Excel が止まるので遅かった)。
  ここはブラウザの読み上げ(Web Speech API。Windows に入っている日本語の声)を使う。画面は止まらない。
  - 前の読み上げが残っていれば止めてから読む(続けて測っても、遅れて溜まらない)
  - 声は最初に 1 回だけ探して覚える(日本語の声。無ければブラウザの既定)
  - ON/OFF はこの PC のブラウザに覚える(読めなくても動く)
*/
const KEY = "cask.read";
let voice = null;

function pickVoice() {
  if (!("speechSynthesis" in window)) return;
  const vs = window.speechSynthesis.getVoices();
  voice = vs.find((v) => v.lang === "ja-JP" && /Haruka|Ayumi|Nanami|Sayaka/i.test(v.name))
    || vs.find((v) => v.lang && v.lang.startsWith("ja")) || null;
}
if ("speechSynthesis" in window) {
  pickVoice();
  window.speechSynthesis.addEventListener?.("voiceschanged", pickVoice);
}

export function canSpeak() { return "speechSynthesis" in window; }

export function speak(text, { onstart, onend } = {}) {
  if (!canSpeak() || !text) return false;
  const s = window.speechSynthesis;
  s.cancel();
  const u = new SpeechSynthesisUtterance(String(text));
  u.lang = "ja-JP";
  if (voice) u.voice = voice;
  u.rate = 1.2;
  if (onstart) u.onstart = onstart;
  if (onend) { u.onend = onend; u.onerror = (e) => onend(e); }
  s.speak(u);
  return true;
}

/** 使う声の名前(設定の「試す」で見せる)。日本語の声が無ければ "" */
export function voiceName() {
  if (!voice) pickVoice();
  return voice ? `${voice.name}(${voice.lang})` : "";
}

export function readOn() {
  try { return localStorage.getItem(KEY) !== "off"; } catch { return true; }
}

export function setReadOn(on) {
  try { localStorage.setItem(KEY, on ? "on" : "off"); } catch { /* 覚えられなくても動く */ }
}
