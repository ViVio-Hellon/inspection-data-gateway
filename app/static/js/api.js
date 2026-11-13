/*
  api.js — サーバとのやりとり

  JSが持つのは「呼ぶ」「受け取ったものを描く」だけ。
  **業務の判断はサーバが行う**(押せるか・隠すか・何と出すか)。
  詳細は docs/設計.md §1。
*/

import * as busy from "./busy.js";
import * as progress from "./progress.js";

const TOKEN = window.APP.token;

/**
 * サーバが返したエラーを、そのまま画面に出せる形で持つ。
 *
 * 断りの返し方は2通りある(設計書 §6.1)。**どちらでも文言はサーバが
 * 持っている**ので、両方から拾う:
 *
 *   入力の形の誤り  … `{"error": {"code", "message"}}`      状態は動いていない
 *   業務としての断り … 画面ぜんぶ + 最上位の `message`       状態は動いていることがある
 *
 * 後者を拾い忘れると、理由(「先にパレットサイズを設定してください。」)の
 * 代わりに「通信に失敗しました (HTTP 422)」と出る ── **通信は成功して
 * いるのに通信の失敗として案内される**ので、現場は直しようがない。
 */
export class ApiError extends Error {
  constructor(status, body, rid = "") {
    const info = (body && body.error) || {};
    super(info.message || (body && body.message)
          || `通信に失敗しました (HTTP ${status})`);
    this.name = "ApiError";
    this.status = status;
    this.code = info.code || "";
    this.field = info.field || "";
    // 断られたときも本文に画面ぜんぶが入っていることがある(422)。
    // 「断られた」と「画面が古いまま」を同時に起こさないために渡す
    this.body = body;
    // 記録番号(`X-Request-Id`)。設定「ログ」でこの操作の前後が引ける
    this.rid = rid || (body && body.rid) || "";
  }
}

/* 2026-11-13: 問い合わせに時間の上限を付ける。Edge で最小化して長く放っておくと、返ってこない問い合わせが
   同じ相手への通り道(6本)を埋め、戻しても「バックエンドに接続できません」のまま戻らなかった。
   上限を過ぎたら打ち切り、切れたと分かったら残っているものもまとめて打ち切る(abortAll)。 */
export const DEFAULT_TIMEOUT_MS = 60000;          // 読むだけ(GET)
export const WRITE_TIMEOUT_MS = 300000;           // 書く操作は長めに待つ(保存・印刷の途中で切らない)
const inflight = new Set();
export function abortAll() {
  // 読むだけの問い合わせだけ打ち切る(書く操作は上限まで待つ)
  for (const c of [...inflight]) {
    if (c.write) continue;
    try { c.abort(); } catch { /* 済み */ }
    inflight.delete(c);
  }
}

async function request(path, options = {}) {
  /*
    **押した本人に、いま待っていることを伝える。**

    どの画面も「押す → サーバへ投げる → 返ってきた画面を描く」で
    出来ているので、投げるところを1か所だけ捕まえれば、画面を
    1つずつ書き換えなくても全部に効く。伝える相手は
    `busy.watchClicks()` が覚えている**いま押されたボタン**。

    250ms より速く返るものには何も出さない ── 一瞬で終わるものに
    待機の姿を出すと、画面がちらつくだけになる。
  */
  const done = busy.mark(busy.pressed());
  const prog = progress.begin(path, TOKEN);           // 少しでも待つときは進み具合を出す(progress.js)
  const ctrl = new AbortController();
  const { timeoutMs, ...opts } = options;
  ctrl.write = !!opts.method && opts.method !== "GET";
  inflight.add(ctrl);
  const timer = setTimeout(() => ctrl.abort(),
                           timeoutMs || (ctrl.write ? WRITE_TIMEOUT_MS : DEFAULT_TIMEOUT_MS));
  try {
    const res = await fetch(path, {
      ...opts,
      signal: ctrl.signal,
      cache: "no-store",
      headers: {
        "X-Tool-Token": TOKEN,
        // どのタブからの要求か。自動終了の見張りが画面ごとに数える
        "X-Screen-Id": window.APP.screenId || "",
        // 同じタブで読み込み直す前の番号(使っている画面を引き継ぐ。screen.js)
        ...(window.APP.screenPrev ? { "X-Screen-Prev": window.APP.screenPrev } : {}),
        // ファイルを送るとき(FormData)は、ブラウザに区切りごと決めさせる
        ...(typeof options.body === "string" ? { "Content-Type": "application/json" } : {}),
        ...(opts.headers || {}),
      },
    });

    let body = null;
    const type = res.headers.get("Content-Type") || "";
    if (type.includes("application/json")) {
      body = await res.json();
    }
    if (!res.ok) {
      const err = new ApiError(res.status, body, res.headers.get("X-Request-Id") || "");
      // ほかの画面が使っている(screen.js が覆いを出す)
      if (err.code === "inactive_screen") document.dispatchEvent(new CustomEvent("app:inactive-screen"));
      throw err;
    }
    return body;
  } finally {
    clearTimeout(timer);
    inflight.delete(ctrl);
    done();
    prog();
  }
}

/*
  `<img src>` や `<iframe src>` のように、**ヘッダを付けられない**読み込みの
  ためのURL。起動トークンはクエリでも受け付けるので、そこに載せる
  (`app/__init__.py` の `request.args.get("t")`)。

  トークンはもともとこのページの `window.APP.token` に入っている。
  ここでURLに付けても新しく漏れるものは無い ── 送信先は自分自身で、
  `Referrer-Policy: no-referrer` により外へ持ち出されることもない。
*/
export function tokenUrl(path) {
  return `${path}${path.includes("?") ? "&" : "?"}t=${encodeURIComponent(TOKEN)}`;
}

export const api = {
  get: (path, opts) => request(path, opts || {}),
  post: (path, data) => request(path, { method: "POST", body: JSON.stringify(data ?? {}) }),
  // ファイルを送る(「表を持ってくる」のドラッグ&ドロップ)
  postForm: (path, form) => request(path, { method: "POST", body: form }),
};
