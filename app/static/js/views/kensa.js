/*
  検査画面(A案「計測コックピット」)。
  判断はすべてサーバ(ufkensa/kensa/session.py)。ここは受け取った view を描き、操作を送るだけ。

  操作の返り値(Result): { view, messages, question, stop }
    messages … 知らせ(VBA の MsgBox の原文)。トーストで出す
    question … 聞くこと。確認窓で答えを聞き、answers に入れて同じ操作を呼び直す
    stop     … 止めた理由。状態は保ったまま(D10)
*/

import { api } from "../api.js";
import { toast, toastError } from "../toast.js";
import { confirmBox } from "../askbox.js";

const $ = (id) => document.getElementById(id);
let V = null;
const recent = [];

const LEVEL_KIND = { info: "info", warn: "warn", error: "error" };

async function op(name, body = {}, answers = {}) {
  let res;
  try {
    res = await api.post(`/api/kensa/${name}`, { ...body, answers });
  } catch (e) {
    toastError(e);
    return null;
  }
  for (const m of res.messages || []) toast(m.text, LEVEL_KIND[m.level] || "info");
  if (res.view) render(res.view);
  if (res.question) {
    const q = res.question;
    const yes = await confirmBox(q.text, { ok: q.buttons === "okcancel" ? "OK" : "はい",
                                           cancel: q.buttons === "okcancel" ? "キャンセル" : "いいえ" });
    return op(name, body, { ...answers, [q.key]: !!yes });
  }
  if (res.stop) toast(res.stop.text, "error");
  return res;
}

function text(el, t) { el.textContent = t ?? ""; }

function kv(dl, pairs) {
  dl.replaceChildren(...pairs.flatMap(([k, v]) => {
    const dt = document.createElement("dt"); dt.textContent = k;
    const dd = document.createElement("dd"); dd.textContent = v ?? "";
    return [dt, dd];
  }));
}

function cellTd(c, cls) {
  const td = document.createElement("td");
  if (cls) td.className = cls;
  if (!c || c.visible === false) { td.dataset.hidden = "1"; return td; }
  td.textContent = c.value ?? "";
  td.dataset.color = c.color || "";
  if (c.border) td.dataset.border = c.border;
  if (c.underline) td.dataset.underline = "1";
  if (c.name) {
    td.dataset.name = c.name;
    td.tabIndex = 0;
    td.title = c.name;
  }
  return td;
}

function renderGrid(g) {
  const t = $("mgrid");
  const head = document.createElement("tr");
  head.append(document.createElement("th"));
  for (const p of g.positions) {
    const th = document.createElement("th");
    th.textContent = p.visible ? p.label : "";
    head.append(th);
  }
  const rows = [head];
  for (const r of g.rows) {
    if (!r.visible) continue;
    const tr = document.createElement("tr");
    const th = document.createElement("th");
    th.className = "rowh";
    th.textContent = r.label;
    tr.append(th, ...r.cells.map((c) => cellTd(c)));
    rows.push(tr);
  }
  for (const [label, list, cls] of [[g.diag_label, g.diag, "sub"], [g.diag_kou_label, g.diag_kou, "sub"],
                                   [g.lateral_label, g.lateral, "sub"], ["判定", g.ng, "sub"]]) {
    if (!list.some((c) => c.visible !== false)) continue;
    const tr = document.createElement("tr");
    tr.className = "sep";
    const th = document.createElement("th");
    th.className = "rowh";
    th.textContent = label;
    tr.append(th, ...list.map((c) => cellTd({ ...c, name: "" }, cls)));
    rows.push(tr);
  }
  t.replaceChildren(...rows);
}

function renderSub(v) {
  const t = $("subgrid");
  const pos = v.grid.positions;
  const head = document.createElement("tr");
  head.append(document.createElement("th"), ...pos.map((p) => {
    const th = document.createElement("th"); th.textContent = p.visible ? p.label : ""; return th;
  }));
  const line = (label, list) => {
    const tr = document.createElement("tr");
    const th = document.createElement("th"); th.className = "rowh"; th.textContent = label;
    tr.append(th, ...list.map((c) => cellTd({ ...c, name: c.visible === false ? "" : `${label}:${c.pos}` }, "sub")));
    return tr;
  };
  t.replaceChildren(head, line("定盤", v.strain.z), line("ゲージ", v.strain.g), line("BH", v.bh), line("不良", v.bad));
  text($("faceLv"), v.strain.face === "Now_B" ? "B面" : "A面");
}

function renderLeft(v) {
  const h = v.header, t = v.tolerance;
  kv($("lotKv"), [["ライン", h.line + (h.line_no || "")], ["Lot", h.lot], ["検査No", h.ken_no], ["用途", h.youto],
                  ["材質/調質", [h.zai, h.tyu].filter(Boolean).join(" ")], ["コメント", h.comment]]);
  kv($("tolKv"), [["厚", t.atu && `${t.atu}  +${t.ap}/-${t.am}  (${t.amk}〜${t.apk})`],
                  ["巾", t.hab && `${t.hab}  +${t.hp}/-${t.hm}  (${t.hmk}〜${t.hpk})`],
                  ["丈", t.tak && `${t.tak}  +${t.tp}/-${t.tm}  (${t.tmk}〜${t.tpk})`],
                  ["対角", t.taik], ["ラテラル", t.lateral_text || t.lateral_grade], ["フリーソリ", t.flatnes_text]]);
  text($("tolNote"), [t.additional, t.conductiv_text].filter(Boolean).join(" / "));
  const works = $("works");
  works.replaceChildren(...h.works.filter((w) => w.visible).map((w) => {
    const b = document.createElement("button");
    b.type = "button"; b.className = "btn"; b.textContent = w.caption;
    b.setAttribute("aria-pressed", String(h.work === w.key));
    b.addEventListener("click", () => op("work", { work: w.key }));
    return b;
  }));
  $("yamaBox").hidden = !h.yama_visible;
  if (h.yama_visible && $("yamaSel").options.length === 0) {
    $("yamaSel").append(...["", "1", "2", "3", "4", "5", "6", "7"].map((y) => new Option(y || "-", y)));
  }
  $("yamaSel").value = h.yama || "";
  for (const [id, b] of [["atBtn", h.at_butt], ["inspeBtn", h.inspe_butt]]) {
    $(id).hidden = !b.visible; $(id).textContent = b.caption;
  }
  text($("atCount"), h.at_count);
  $("proofLv").textContent = h.proofread ? "校正済" : "未校正";
  $("proofLv").className = "lv " + (h.proofread ? "lv--ok" : "lv--ng");
  const workers = $("workers");
  workers.replaceChildren(...(v.workers || []).map((w) => { const li = document.createElement("li"); li.textContent = w; return li; }));
}

function renderRight(v) {
  text($("nextName"), v.next.label || v.next.name || "-");
  text($("nextHint"), [v.next.machin, v.next.auto ? "" : "手で選んだ欄"].filter(Boolean).join(" / "));
  const m = v.minmax;
  kv($("mmKv"), [["厚", m.amin && `${m.amin} 〜 ${m.amax}`], ["巾", m.hmin && `${m.hmin} 〜 ${m.hmax}`], ["丈", m.tmin && `${m.tmin} 〜 ${m.tmax}`]]);
  if (v.checks) {
    $("checks").replaceChildren(...v.checks.map((c) => {
      const li = document.createElement("li"); li.textContent = c.label; li.dataset.ok = c.ok === null ? "" : (c.ok ? "1" : "0"); return li;
    }));
  }
  $("blind").hidden = !v.blind;
  text($("blind"), v.blind);
  text($("workLv"), [v.header.work, v.header.yama && `${v.header.yama}山`].filter(Boolean).join(" "));
}

function render(v) {
  V = v;
  renderLeft(v);
  renderGrid(v.grid);
  renderSub(v);
  renderRight(v);
}

function remember(line) {
  recent.unshift(`${new Date().toLocaleTimeString("ja-JP")} ${line}`);
  recent.length = Math.min(recent.length, 8);
  $("recent").replaceChildren(...recent.map((l) => { const li = document.createElement("li"); li.textContent = l; return li; }));
}

function onEnter(id, fn) {
  $(id).addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    const val = $(id).value;
    $(id).value = "";
    if (val) fn(val);
  });
}

function init() {
  onEnter("cardIn", (card) => op("worker", { card }));
  onEnter("lotIn", (lot) => op("lot", { lot }));
  onEnter("measIn", async (t) => { remember(t); await op("receive", { text: t }); $("measIn").focus(); });
  $("yamaSel").addEventListener("change", () => op("yama", { yama: $("yamaSel").value }));
  $("atBtn").addEventListener("click", () => op("at"));
  $("inspeBtn").addEventListener("click", () => op("inspe"));
  $("mmBtn").addEventListener("click", () => op("minmax"));
  $("printBtn").addEventListener("click", async () => {
    const r = await op("print-check");
    if (r && !r.stop && r.view && r.view.print_ready) await op("print-commit");
  });
  $("newBtn").addEventListener("click", () => op("new"));
  $("mgrid").addEventListener("click", (e) => {
    const td = e.target.closest("td[data-name]");
    if (td) { op("click", { name: td.dataset.name }); $("measIn").focus(); }
  });
  $("mgrid").addEventListener("dblclick", async (e) => {
    const td = e.target.closest("td[data-name]");
    if (!td || !td.textContent) return;
    op("dblclick", { name: td.dataset.name });
  });
  api.get("/api/kensa/state").then((res) => { render(res.view); return op("start"); }).catch(toastError);
  $("measIn").focus();
}

init();
