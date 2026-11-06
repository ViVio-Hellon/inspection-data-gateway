/* 受信機(U-WAVE-R・DDM-100)から直接受け取る値を、いま開いている画面へ渡す(2026-11-06)。

   サーバは届いた値を「最後に取りに来た画面」にだけ渡す。開いていない画面には渡さない
   (前は通常検査にだけ入っていたので、校正などで測った値が見えないまま検査に入るおそれがあった)。
   画面が隠れている(別タブ・最小化)間は取りに行かないので、その画面には入らない。
   使い方: startReceiver("calibration", (text) => op("receive", { text }), (st) => 表示を変える) */
import { api } from "./api.js";

export async function startReceiver(screen, deliver, onState = () => {}) {
  let last = 0;
  let on = false;
  try {
    const d = await api.post("/api/kensa/recv/start", {});
    on = !!d.on;
    onState({ on, port: d.port, error: d.error });
    if (!on) return;
    const p = await api.get(`/api/kensa/recv?since=999999999&screen=${screen}`);
    last = p.last;                                      // 開く前に届いた分は流さない
  } catch (e) { return; }
  const tick = async () => {
    const visible = document.visibilityState === "visible";
    try {
      const d = await api.get(`/api/kensa/recv?since=${last}&screen=${screen}&visible=${visible ? 1 : 0}`);
      for (const r of d.results) await deliver(r.text, r);
      last = d.last;
      onState({ on: d.running, port: d.port, error: d.error });
    } catch (e) { /* 次の回で */ }
    setTimeout(tick, 400);
  };
  tick();
}
