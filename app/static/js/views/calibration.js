/*
  校正画面(VBA の UF校正)。
  判断はすべてサーバ(ufkensa/kensa/calibration.py → calibration_session.py)。ここは view を描いて操作を送るだけ。
  返り値: { view, messages, question, stop }(検査画面と同じ形。question は使わない)
*/

import { api } from "../api.js";
import { toast, toastError } from "../toast.js";
import { playMessage } from "../sound.js";
import { renderRecvStatus, startReceiver } from "../recv.js";

const $ = (id) => document.getElementById(id);
const LEVEL_KIND = { info: "info", warn: "warn", error: "error" };
let V = null;
// 2026-11-10: 検査員の登録から移ってきた(校正が済んでいない)ときの案内を出す
const fromWorker = new URLSearchParams(location.search).get("from") === "worker";

function text(el, t) { el.textContent = t ?? ""; }

async function op(name, body = {}) {
  let res;
  try {
    res = await api.post(`/api/calibration/${name}`, body);
  } catch (e) {
    toastError(e);
    return null;
  }
  for (const m of res.messages || []) {
    playMessage(m);                                               // 音(D29: 出来事の鍵。鳴らせなければ文言)
    if (m.text) toast(m.text, LEVEL_KIND[m.level] || "info");
  }
  if (res.view) render(res.view);
  if (res.stop) toast(res.stop.text, "error");
  if (res.view && res.view.closed) backToKensa();
  return res;
}

/* 校正に合格(F1 = "end")したら、検査の判定をやり直して検査画面へ戻る(UF校正.INPG_Exit) */
let leaving = false;
async function backToKensa() {
  if (leaving) return;
  leaving = true;
  try { await api.post("/api/kensa/calibrated", { answers: {} }); } catch (e) { /* 検査画面が開いていなければ何もしない */ }
  toast("校正が終わりました。検査画面へ戻ります", "success");
  setTimeout(() => { location.href = "/kensa"; }, 1200);
}

function renderStatus(v) {
  const s = v.status;
  const box = $("calStatus");
  if (!s) {
    box.dataset.ok = "0";
    text($("calBadge"), "!");
    text($("calStatusText"), "ラインが設定されていません");
    text($("calStatusSub"), "設定画面でこの端末のラインを決めてください");
    $("calStatusKv").replaceChildren();
    return;
  }
  box.dataset.ok = s.calibrated ? "1" : "0";
  $("calFrom").hidden = !(fromWorker && !s.calibrated);
  text($("calBadge"), s.calibrated ? "済" : "未");
  text($("calStatusText"), s.calibrated ? `${s.date} ${s.shift} の校正は済んでいます` : `${s.date} ${s.shift} の校正が済んでいません`);
  text($("calStatusSub"), s.skip_text || (s.calibrated ? "検査画面で測定できます" : "校正を記録するまで、検査画面の測定値は受け付けません"));
  const pairs = [["いま", s.now], ["最後の記録", [s.last_date, s.last_shift].filter(Boolean).join(" ") || "なし"]];
  const se = v.straight_edge;
  if (se) pairs.push(["ストレートエッヂ", `${se.lo} 〜 ${se.hi}` + (se.length ? `(長さ ${se.length} ±${se.tol})` : "")]);
  for (const m of v.masters || []) pairs.push([`参照マスタ(${m.label})`, `${m.path}　表「${m.table}」`]);
  $("calStatusKv").replaceChildren(...pairs.flatMap(([k, x]) => {
    const dt = document.createElement("dt"); dt.textContent = k;
    const dd = document.createElement("dd"); dd.textContent = x;
    return [dt, dd];
  }));
}

function renderStep(v) {
  const step = v.step || "1000";
  $("calStep").dataset.step = step;
  for (const li of document.querySelectorAll("#calStep .cal__flow li")) {
    const order = ["1000", "gauge", "done"];
    li.dataset.state = li.dataset.k === step ? "now" : (order.indexOf(li.dataset.k) < order.indexOf(step) ? "past" : "");
  }
  text($("calGauge"), v.gauge || "-");
  text($("calValue"), v.sokutei || "-");
  const se = v.straight_edge;
  text($("calRange"), step === "1000" ? `${v.range_1000[0]} 〜 ${v.range_1000[1]}`
    : (se ? `${se.lo} 〜 ${se.hi}` : "校正ストレートエッヂ に このラインの行がありません"));
  $("calNg").hidden = !v.ng_value;
  text($("calNg"), v.ng_value);
  $("calSaiken").hidden = !v.saiken;
  $("calDone").hidden = step !== "done";
  text($("calLine"), v.line || "-");
  text($("calKname"), v.kname || "(検査員の登録なし)");
  $("calSkip").hidden = !v.can_skip_1000;
  $("calIn").disabled = step === "done";
  text($("calNote"), v.accepting === false && step !== "done" ? "ｸﾘｱ＆ﾘｾｯﾄ を押すと、次の測定値を受け付けます。"
    : (v.template_created ? "品証検証 シートが無かったので、VBA のブックと同じ枠を作りました。" : ""));
}

function renderRanks(v) {
  text($("calMonth"), v.month);
  text($("calRankKind"), v.rank_kind ? `${v.rank_kind}の順位` : "");
  text($("calCenter"), v.center ? `このラインの中心値: ${v.center}(以上はﾌﾟﾗｽ外れ、未満はﾏｲﾅｽ外れ)` : "このラインの中心値はありません");
  const li = (t) => { const e = document.createElement("li"); e.textContent = t || "-"; return e; };
  const ranks = v.ranks || [];
  $("calRankMax").replaceChildren(...ranks.filter((r) => r.pos <= 3).map((r) => li(r.text)));
  $("calRankMin").replaceChildren(...ranks.filter((r) => r.pos > 3).map((r) => li(r.text)));
}

function renderRecords(v) {
  text($("calRecCount"), `${v.records.length} 件`);
  $("calRecords").tBodies[0].replaceChildren(...v.records.map((r) => {
    const tr = document.createElement("tr");
    tr.dataset.ng = r.ng ? "1" : "";
    for (const t of [r.date + (r.night ? " (夜勤)" : ""), r.shift, r.time, r.value, r.kname, r.line, r.ng ? "不合格" : "合格"]) {
      const td = document.createElement("td"); td.textContent = t; tr.append(td);
    }
    return tr;
  }));
}

function render(v) {
  V = v;
  renderStatus(v);
  renderStep(v);
  renderRanks(v);
  renderRecords(v);
  focusIn();
}

/** 測定器の入力を止めない: 受信欄にいつもフォーカスを戻す(D6)。 */
function focusIn() {
  const el = $("calIn");
  if (!el.disabled && document.activeElement !== el) el.focus();
}

$("calIn").addEventListener("keydown", async (e) => {
  if (e.key !== "Enter" && !(e.key === "Tab" && $("calIn").value)) return;   // VBA の AfterUpdate(Enter / Tab)
  e.preventDefault();
  const t = e.target.value;
  e.target.value = "";
  if (t.trim() === "") return;
  await op("receive", { text: t });
});
$("calClear").addEventListener("click", () => op("clear"));
$("calSkip").addEventListener("click", () => op("skip-1000"));
$("calRestart").addEventListener("click", () => op("open"));
$("calRankH").addEventListener("click", () => op("rank", { kind: "hinsho" }));
$("calRankK").addEventListener("click", () => op("rank", { kind: "kousei" }));
// 画面の中(#cal)だけで拾う。document に付けると画面を移るたびに重なる(nav.js は中身を差し替えるだけ)
$("cal").addEventListener("click", (e) => {
  if (!(e.target instanceof Element) || !e.target.closest("input, select, textarea, a")) setTimeout(focusIn, 0);
});

op("open").catch(toastError);

// 受信機から直接受け取る(保守で「検査画面で受け取る」にした受信機)。この画面が開いている間だけ、この画面に入る
// 2026-11-09: 入らなかった・合格しなかった理由を受信状態の行に出す
let calAccepting = true;
startReceiver("calibration", async (text) => {
  const was = calAccepting;
  const res = await op("receive", { text });
  const v = res && res.view;
  if (v) calAccepting = v.accepting !== false;
  if (!res || res.stop) return res;
  if (!was) return { stop: { text: "ｸﾘｱ＆ﾘｾｯﾄ を押すまで次の測定値を受け付けません" } };
  if (v && v.ng_value) return { stop: { text: v.ng_value + (v.saiken ? "(ｸﾘｱを押してもう1度測定してください)" : "") } };
  return res;
}, (st) => renderRecvStatus($("calRecv"), st));
