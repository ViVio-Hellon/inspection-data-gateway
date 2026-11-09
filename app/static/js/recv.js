/* 受信機(U-WAVE-R・DDM-100)から直接受け取る値を、いま開いている画面へ渡す(2026-11-06)。

   サーバは届いた値を「最後に取りに来た画面」にだけ渡す。開いていない画面には渡さない
   (前は通常検査にだけ入っていたので、校正などで測った値が見えないまま検査に入るおそれがあった)。
   画面が隠れている(別タブ・最小化)間は取りに行かないので、その画面には入らない。
   使い方: startReceiver("calibration", (text) => op("receive", { text }), (st) => 表示を変える) */
import { api } from "./api.js";

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
