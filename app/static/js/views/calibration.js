/*
  校正画面(VBA の UF校正)。
  判断はすべてサーバ(ufkensa/kensa/calibration.py → calibration_session.py)。ここは view を描いて操作を送るだけ。
  返り値: { view, messages, question, stop }(検査画面と同じ形。question は使わない)
*/

import { api, tokenUrl } from "../api.js";
import { toast, toastError } from "../toast.js";

const $ = (id) => document.getElementById(id);
const LEVEL_KIND = { info: "info", warn: "warn", error: "error" };
let V = null;

function text(el, t) { el.textContent = t ?? ""; }

/** VBA の音声(合格 chime2.wav / 不合格 tada.wav)。無ければ鳴らないだけ。 */
function play(name) {
  try {
    const a = new Audio(tokenUrl(`/api/kensa/sound/${encodeURIComponent(name)}`));
    a.play().catch(() => {});
  } catch (e) { /* 鳴らせなくても止めない */ }
}

async function op(name, body = {}) {
  let res;
  try {
    res = await api.post(`/api/calibration/${name}`, body);
  } catch (e) {
    toastError(e);
    return null;
  }
  for (const m of res.messages || []) {
    if (m.sound) play(m.sound);
    if (m.text) toast(m.text, LEVEL_KIND[m.level] || "info");
  }
  if (res.view) render(res.view);
  if (res.stop) toast(res.stop.text, "error");
  return res;
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
  text($("calBadge"), s.calibrated ? "済" : "未");
  text($("calStatusText"), s.calibrated ? `${s.date} ${s.shift} の校正は済んでいます` : `${s.date} ${s.shift} の校正が済んでいません`);
  text($("calStatusSub"), s.skip_text || (s.calibrated ? "検査画面で測定できます" : "校正を記録するまで、検査画面の測定値は受け付けません"));
  const pairs = [["いま", s.now], ["最後の記録", [s.last_date, s.last_shift].filter(Boolean).join(" ") || "なし"]];
  const se = v.straight_edge;
  if (se) pairs.push(["ストレートエッヂ", `${se.lo} 〜 ${se.hi}` + (se.length ? `(長さ ${se.length} ±${se.tol})` : "")]);
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
  if (e.key !== "Enter") return;
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
