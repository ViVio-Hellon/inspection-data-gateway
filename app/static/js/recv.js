/* 受信機(U-WAVE-R・DDM-100)から直接受け取る値を、いま開いている画面へ渡す(2026-11-06)。

   サーバは届いた値を「最後に取りに来た画面」にだけ渡す。開いていない画面には渡さない
   (前は通常検査にだけ入っていたので、校正などで測った値が見えないまま検査に入るおそれがあった)。
   画面が隠れている(別タブ・最小化)間は取りに行かないので、その画面には入らない。
   使い方: startReceiver("calibration", (text) => op("receive", { text }), (st) => 表示を変える)
   検査画面は、受信機の値をサーバが検査のセッションに渡している(カーソルに関係なく入る)。
   画面には実行した結果が届くので、deliver に fromServer(op) を渡して、操作の結果と同じ手順で出す */
import { api } from "./api.js";
import { toast } from "./toast.js";

/* 2026-11-09: 「検査できない理由を表示して」(現場の声)。受信機の状態・最後に届いた値と時刻・入らなかった理由。
   st = startReceiver の onState に来るもの。el に 1 行(と理由の行)を描く */
export function renderRecvStatus(el, st) {
  if (!el) return;
  const dev = st.on ? `受信機 ${st.port}: 受信中` : `受信機: 止まっている — ${st.why || st.error || "理由不明"}`;
  const last = st.received && st.received.text
    ? `最後に届いた値 ${st.received.text}(${st.received.at})` : "まだ値は届いていません";
  el.dataset.on = st.on ? "1" : "";
  el.replaceChildren();
  const a = document.createElement("span"); a.className = "recv-status__dev"; a.textContent = dev;
  const b = document.createElement("span"); b.className = "recv-status__last"; b.textContent = last;
  el.append(a, b);
  if (st.rejected) {
    const c = document.createElement("span"); c.className = "recv-status__why"; c.setAttribute("role", "alert");
    c.textContent = `入らなかった理由: ${st.rejected}`;
    el.append(c);
  }
}

/* 画面の操作の結果(res)から、届いた値が入らなかった理由を拾う(止め・警告)。入ったなら "" */
export function rejectReason(res) {
  if (!res) return "受け取れませんでした(通信エラー)";
  if (res.stop) return res.stop.text;
  const w = (res.messages || []).find((m) => m.level === "warn" || m.level === "error");
  return w ? w.text : "";
}

export async function startReceiver(screen, deliver, onState = () => {}) {
  let last = 0;
  let st = { on: false };
  const show = (d) => { st = { ...st, on: !!(d.running ?? d.on), port: d.port, error: d.error, why: d.why,
                               received: d.received || d.last || st.received }; onState(st); };
  try {
    const d = await api.post("/api/kensa/recv/start", {});
    show(d);
    const p = await api.get(`/api/kensa/recv?since=999999999&screen=${screen}`);
    last = p.last;                                      // 開く前に届いた分は流さない
  } catch (e) { return; }
  const tick = async () => {
    const visible = document.visibilityState === "visible";
    try {
      const d = await api.get(`/api/kensa/recv?since=${last}&screen=${screen}&visible=${visible ? 1 : 0}`);
      for (const r of d.results) {
        const res = await deliver(r.text, r);
        // 2026-11-09: 入らなかった理由を残す(状態の行に出す)
        st.rejected = res === undefined ? "" : rejectReason(res);
      }
      last = d.last;
      show(d);
    } catch (e) { /* 次の回で */ }
    setTimeout(tick, st.on ? 400 : 3000);               // 止まっていても状態(理由)は見続ける
  };
  tick();
}

/* サーバが受け取って実行した結果 r(view・messages・question・stop・why)を、画面の操作の結果として出す。
   op = screen_op.js の screenOp(...)。問いを聞かれたら、答えを付けて同じ値を送り直す(op が okcancel も扱う)。
   2026-11-09: 入らなかった理由を見える所に出す(知らせ・止めで出ていないときもトーストで出す)。
   before(r) … 出す前に画面ごとにすること(最近の入力に残す など) */
export function fromServer(op, before = () => {}) {
  return async (text, r) => {
    before(r);
    const res = await op.handle(r, "receive", { text });
    if (!r.stop && r.why && !(r.messages || []).some((m) => m.text === r.why)) {
      toast(`受信した値 ${text} は入りませんでした: ${r.why}`, "warn");
    }
    return res;
  };
}
