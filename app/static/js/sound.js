/*
  sound.js — 音を鳴らす(どの画面もここを使う。D29、all-tools 日報ツールと同じ決まり)

  **鳴らすのはここ、決めるのはサーバ。** 業務の部品は知らせ(Message)の `sound` に出来事の鍵
  (`kensa_ok` など)を入れて返す。どのファイルを鳴らすか・鳴らさないかは 設定 →「音」で決まり、
  サーバが `/api/sound/<出来事>` で配る(起動トークン付き)。ここはファイル名を持たない。

  【鳴らしてよい出来事の一覧】
  `/api/sound/cues`(鳴らすと決めてあり、ファイルがある出来事)を開いたときと、画面へ戻ったときに読む。
  一覧に無い出来事は鳴らしに行かない(ファイルが無い音を毎回取りに行かない)。読めないあいだは鳴らしに行く。

  【ブラウザは勝手に音を鳴らせない】
  利用者がそのページを一度も触っていないあいだ、ブラウザは play() を断る(自動再生の制限)。
    ・最初のクリック/キー入力で無音を1回鳴らして「触った」ことにする
    ・それでも鳴らせなかったら、**画面に文言を出す**(音は補助で、伝えたいことは文字のほうにある)
*/

import { tokenUrl } from "./api.js";
import { toast } from "./toast.js";

let cues = null;          // Set(鳴らせる出来事)。読めるまでは null
let unlocked = false;
let hinted = false;

/** 鳴らしてよい出来事を読み直す。 */
export async function loadCues() {
  try {
    const res = await fetch("/api/sound/cues", {
      cache: "no-store",
      headers: { "X-Tool-Token": window.APP.token, "X-Screen-Id": window.APP.screenId || "" },
    });
    if (!res.ok) return;
    const body = await res.json();
    if (Array.isArray(body.cues)) cues = new Set(body.cues);
  } catch (e) { /* 読めなければ前のまま */ }
}

function unlock() {
  if (unlocked) return;
  unlocked = true;
  try {
    const el = new Audio();
    // 1サンプルぶんの無音(WAV)。外部ファイルを読まない(音声フォルダの設定に関係なく動く)
    el.src = "data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAgD4AAAB9AAACABAAZGF0YQAAAAA=";
    el.volume = 0;
    el.play().catch(() => {});
  } catch (e) { /* 音が使えない端末。文言だけで動く */ }
}

async function playUrl(url) {
  try {
    // 1回ごとに新しく作る(VBA の mciSendString "Play" と同じく、続けて来た音を重ねて鳴らす)
    const a = new Audio(url);
    await a.play();
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * 出来事の音を鳴らす。鳴らせなければ `text`(サーバが付けた文言)を出す。
 * 鳴らさないと決めてある・ファイルが無い出来事は何もしない(文言は画面の本体に出ている)。
 */
export async function playKey(key, text = "") {
  if (!key) return false;
  if (cues && !cues.has(key)) return false;
  const ok = await playUrl(tokenUrl(`/api/sound/${encodeURIComponent(key)}`));
  if (!ok) {
    if (text) toast(text, "info");
    if (!hinted) {
      hinted = true;
      console.info("[sound] 鳴らせませんでした(画面を一度押すと鳴るようになります)。文言だけ出します:", key);
    }
  }
  return ok;
}

/** サーバの知らせ(`{text, sound, sound_text}`)の音を鳴らす。 */
export function playMessage(m) {
  if (m && m.sound) playKey(m.sound, m.sound_text || "").catch(() => {});
}

/** 設定画面の試聴。選んであるファイル(保存前でも)を名前で鳴らす。音声フォルダの中だけ。 */
export async function previewFile(name) {
  unlock();
  return playUrl(tokenUrl(`/api/sound/preview?name=${encodeURIComponent(name)}`));
}

for (const type of ["pointerdown", "keydown"]) {
  document.addEventListener(type, unlock, { once: true, passive: true, capture: true });
}
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") loadCues();
});
loadCues();
