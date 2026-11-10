/* 画面の操作をサーバへ送る共通の手順(検査・キャスク・全ヵ所・Aインプット・校正。整理 6、2026-11-10)。

   前は画面ごとに op() を写していて、はい/いいえの扱い・止めの出し方が少しずつ違っていた。
   操作の返り値(Result): { view, messages, question, stop }
     1. 送る(POST /api/<画面>/<操作>  本文 = { ...body, answers })。通信エラーはトーストで出して null
     2. messages … 1 件ずつ showMessage(既定: 音 + トースト)
     3. view     … render
     4. question … 確認窓(okcancel なら OK/キャンセル、ほかは はい/いいえ)で答えを聞き、answers に入れて同じ操作を呼び直す
                   (askInput があれば、文字を打ってもらう問い(buttons = "input")をそちらで聞く)
     5. stop     … トースト(既定は error。stopKind で変えられる)
     6. after(name, res, opts) … 画面ごとの後始末(焦点を戻す・帳票の案内 など)

   使い方:
     const op = screenOp("cask", { render, after: (name, res, o) => { if (o.refocus !== false) focusInput(); } });
     await op("lot", { lot: "H7345B0" });
     op.handle(res, "receive", { text })   // サーバがもう実行した結果(受信機の値)を同じ手順で出す */
import { api } from "./api.js";
import { toast, toastError } from "./toast.js";
import { playMessage } from "./sound.js";
import { confirmBox } from "./askbox.js";

export const LEVEL_KIND = { info: "info", warn: "warn", error: "error" };

/** 既定の知らせの出し方: 音(D29: 出来事の鍵。鳴らせなければ文言)と、文言があればトースト */
export function defaultMessage(m) {
  playMessage(m);
  if (m.text) toast(m.text, LEVEL_KIND[m.level] || "info");
}

/** 問いの答え(はい/いいえ・OK/キャンセル)を聞く */
export function askYesNo(q) {
  const okc = q.buttons === "okcancel";
  return confirmBox(q.text, { ok: okc ? "OK" : "はい", cancel: okc ? "キャンセル" : "いいえ" });
}

export function screenOp(screen, {
  render = () => {},
  showMessage = defaultMessage,
  stopKind = () => "error",
  askInput = null,               // (q) => Promise<本文に足すもの | undefined>(undefined なら はい/いいえ で聞く)
  after = () => {},
} = {}) {
  async function handle(res, name, body = {}, answers = {}, opts = {}) {
    for (const m of res.messages || []) showMessage(m);
    if (res.view) render(res.view);
    if (res.question) {
      const q = res.question;
      if (askInput) {
        const extra = await askInput(q);
        if (extra !== undefined) return op(name, { ...body, ...extra }, answers, opts);
      }
      const yes = await askYesNo(q);
      return op(name, body, { ...answers, [q.key]: !!yes }, opts);
    }
    if (res.stop) toast(res.stop.text, stopKind(res.stop));
    after(name, res, opts);
    return res;
  }
  async function op(name, body = {}, answers = {}, opts = {}) {
    let res;
    try {
      res = await api.post(`/api/${screen}/${name}`, { ...body, answers });
    } catch (e) {
      toastError(e);
      return null;
    }
    return handle(res, name, body, answers, opts);
  }
  op.handle = handle;
  return op;
}
