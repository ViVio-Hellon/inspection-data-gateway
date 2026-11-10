/*
  全ヵ所検査(一般。5N・三菱受託は使っていないので切り替えに出さない D21)の画面。
  判断はすべてサーバ(ufkensa/kensa/zenkasho_session.py)。ここは受け取った view を描き、操作を送るだけ。

  操作の返り値(Result): { view, messages, question, stop }
    messages … 知らせ(VBA の MsgBox の原文)。トーストで出す。音だけのものは鳴らすだけ
    question … 聞くこと。確認窓で答えを聞き、answers に入れて同じ操作を呼び直す。
               buttons = "input" は InputBox(打った文字を choice_input に入れて呼び直す)
    stop     … 止めた理由。状態は保ったまま(D10)
*/

import { api, tokenUrl } from "../api.js";
import { toast, toastError } from "../toast.js";
import { playMessage } from "../sound.js";
import { confirmBox, promptBox } from "../askbox.js";
import { packDialog } from "./pack_dialog.js";
import { startReceiver } from "../recv.js";

const $ = (id) => document.getElementById(id);
const LEVEL_KIND = { info: "info", warn: "warn", error: "error" };
const recent = [];
let V = null;
let lastActive = "";
let flashTimer = 0;
let pack = null;            // 包装仕様書の窓(UF_PackNo)

async function op(name, body = {}, answers = {}) {
  let res;
  try {
    res = await api.post(`/api/zenkasho/${name}`, { ...body, answers });
  } catch (e) {
    toastError(e);
    return null;
  }
  for (const m of res.messages || []) {
    playMessage(m);                                               // 音(D29: 出来事の鍵。鳴らせなければ文言)
    if (m.text) toast(m.text, LEVEL_KIND[m.level] || "info");
  }
  if (res.view) render(res.view);
  if (res.question) {
    const q = res.question;
    if (q.buttons === "input") {
      // 作業選択は「一般」しか使わない(D21)ので、最初から入れておく(Enter だけで開ける。空のまま閉じると VBA どおり止まる)
      const v = await promptBox(q.text, { value: q.key === "mit_choice" ? "一般" : "" });
      return op(name, { ...body, choice_input: v ?? "" }, answers);
    }
    const yes = await confirmBox(q.text, { ok: q.buttons === "okcancel" ? "OK" : "はい",
                                           cancel: q.buttons === "okcancel" ? "キャンセル" : "いいえ" });
    return op(name, body, { ...answers, [q.key]: !!yes });
  }
  if (res.stop) toast(res.stop.text, "error");
  focusMeas();
  return res;
}

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
}

function focusMeas() {
  // 画面は入れ替え式(nav.js)なので、ほかの画面へ移ったあとも document の click から呼ばれる。入力欄が無ければ何もしない
  if (!$("measIn")) return;
  const a = document.activeElement;
  if (document.querySelector("dialog[open]")) return;
  if (!a || !a.matches("input, select, textarea") || a === $("measIn")) $("measIn").focus();
}

// --- 上: 作業の切替 ---------------------------------------------------
function renderTop(v) {
  $("modes").replaceChildren(...v.modes.map((m) => {
    const b = el("button", "zk-mode", m.label);
    b.type = "button";
    b.setAttribute("role", "tab");
    b.setAttribute("aria-selected", String(m.active));
    b.dataset.mode = m.key;
    return b;
  }));
  $("lineName").textContent = v.line ? `ライン ${v.line}` : "";
  $("notice").hidden = !v.notice;
  $("notice").textContent = v.notice;
  $("notice").title = (v.suspect || []).join("\n");
  $("stopped").hidden = v.events_enabled;
  document.body.dataset.zkMode = v.mode || "";
}

// --- 左: フォーム -----------------------------------------------------
function setInput(id, value) {
  const i = $(id);
  if (document.activeElement !== i) i.value = value ?? "";
}

function renderForm(v) {
  const f = v.form;
  setInput("fLot", f.lot); setInput("fKenNo", f.ken_no); setInput("fYouto", f.youto); setInput("fKname", f.kname);
  setInput("fMin", f.min); setInput("fMax", f.max);
  for (const [id, key] of [["fLot", "lot"], ["fKenNo", "ken_no"], ["fYouto", "youto"], ["fKname", "kname"],
                           ["fMin", "min"], ["fMax", "max"]]) $(id).dataset.name = f.names[key];
  $("methods").replaceChildren(...f.methods.map((m) => {
    const b = el("button", "zk-seg__b", m);
    b.type = "button"; b.setAttribute("role", "radio"); b.setAttribute("aria-checked", String(f.method === m));
    b.addEventListener("click", () => op("form", { name: f.names.method, value: m }));
    return b;
  }));
  $("yama").replaceChildren(...f.yama_list.map((y) => {
    const b = el("button", "zk-seg__b", y === "枚数" ? "枚数" : `${y}山`);
    b.type = "button"; b.setAttribute("role", "radio"); b.setAttribute("aria-checked", String(f.yama === y));
    b.addEventListener("click", () => op("form", { name: "YamaLB", value: y }));
    return b;
  }));
  renderKenh(f.kenh.filter((k) => k.visible));
  $("formula").textContent = f.takasa_formula ? `高さ計算式: ${f.takasa_formula}` : "";
  $("ainputLink").hidden = !v.ainput;
  if (v.ainput) {
    $("ainputLink").textContent = `Aインプットへ(${v.ainput.kind} ${v.ainput.method})`;
    $("ainputLink").href = `/ainput?from=${encodeURIComponent(v.ainput.kind)}&method=${encodeURIComponent(v.ainput.method)}`;
  }
}

// F5(2026-11-10): 検入高さの欄は、並び(番号・見出し)が変わったときだけ作り直す。打っている途中の欄は値も触らない
function renderKenh(list) {
  const box = $("kenh");
  const sig = JSON.stringify(list.map((k) => [k.n, k.caption]));
  const typing = box.contains(document.activeElement);
  if (box.dataset.sig !== sig && !typing) {
    box.replaceChildren(...list.map((k) => {
      const lab = el("label", "", k.n === 7 ? "枚数" : `${k.caption || k.n}`);
      const i = el("input", "input");
      i.dataset.name = `KenHBox${k.n}`; i.autocomplete = "off";
      i.addEventListener("change", () => op("form", { name: i.dataset.name, value: i.value }));
      lab.append(i, el("small", "muted", k.n === 7 ? "枚" : "mm"));
      return lab;
    }));
    box.dataset.sig = sig;
  }
  for (const k of list) {
    const i = box.querySelector(`input[data-name="KenHBox${k.n}"]`);
    if (i && document.activeElement !== i) i.value = k.value ?? "";
  }
}

// --- 中央: シートの上部と表 ----------------------------------------------
function renderHead(v) {
  $("head").replaceChildren(...v.header.map((h) => {
    const d = el("div", "zk-head__i");
    d.append(el("dt", "", h.label));
    const dd = el("dd", h.red ? "red" : "", h.text || "-");
    if (h.cell) {
      d.dataset.r = h.cell[0]; d.dataset.c = h.cell[1];
      if (h.edit) d.classList.add("edit");
      if (v.active && v.active.r === h.cell[0] && v.active.c === h.cell[1]) d.classList.add("cur");
    }
    d.append(dd);
    return d;
  }));
}

function renderTable(v) {
  const t = $("sheet");
  const thead = el("thead");
  const g = el("tr", "grp");
  g.append(el("th", "no", "No"));
  for (const x of v.groups) {
    const th = el("th", `g g--${x.label}`);
    th.colSpan = x.span;
    th.append(el("b", "", x.label));
    if (x.tol) th.append(el("small", "", x.tol));
    g.append(th);
  }
  const h = el("tr", "lab");
  h.append(el("th", "no", ""));
  for (const c of v.columns) h.append(el("th", `k-${c.kind}`, c.label));
  thead.append(g, h);
  const tbody = el("tbody");
  const ac = v.active;
  for (const row of v.rows) {
    const tr = el("tr");
    if (ac && ac.r === row.r) tr.className = "currow";
    tr.append(el("th", "no", row.no));
    row.cells.forEach(([text, cls], i) => {
      const col = v.columns[i];
      const td = el("td", `k-${col.kind} ${cls || ""}`.trim(), text);
      td.dataset.r = row.r; td.dataset.c = col.c;
      if (ac && ac.r === row.r && ac.c === col.c) td.classList.add("cur");
      tr.append(td);
    });
    tbody.append(tr);
  }
  t.replaceChildren(thead, tbody);
  const key = ac ? `${v.page}:${ac.r}:${ac.c}` : "";
  if (key && key !== lastActive) {
    const cur = t.querySelector("td.cur");
    if (cur) cur.scrollIntoView({ block: "nearest", inline: "nearest" });
  }
  lastActive = key;
}

function renderFlash(v) {
  if (!v.flash && !v.diag) return;
  const f = $("flash");
  // UF対角: VBA は絵を 2 秒出す(UFOK/UFNG は閉じる)
  f.textContent = v.diag ? "対角" : (v.flash === "OK" ? "○" : "×");
  f.dataset.kind = v.diag ? "diag" : v.flash;
  f.hidden = false;
  clearTimeout(flashTimer);
  flashTimer = setTimeout(() => { f.hidden = true; }, v.diag ? 2000 : 700);
}

// --- 右: 次に入力する欄・操作 --------------------------------------------
function renderRight(v) {
  $("nextName").textContent = v.active ? v.active.label : "-";
  $("nextHint").textContent = v.open ? (v.events_enabled ? "" : "自動判定が止まっています") : "画面を開いていません";
  $("emptyBtn").hidden = !v.can_next_empty;
  $("prints").replaceChildren(...v.prints.map((p) => {
    const b = el("button", "btn btn--print", p.label);
    b.type = "button";
    b.addEventListener("click", () => op("print", { which: p.which }));
    return b;
  }));
  if (v.reports.length) {
    $("reports").replaceChildren(...v.reports.map((r) => {
      const li = el("li");
      const a = el("a", "", `帳票を開く: ${r.label}`);
      a.href = tokenUrl(`/api/zenkasho/report/${encodeURIComponent(r.name)}`);
      a.target = "_blank"; a.rel = "noopener";
      li.append(a);
      return li;
    }), ...(v.saved.length ? [el("li", "muted", `保存 ${v.saved.length} ファイル: ${v.saved[0]}`)] : []));
    const first = v.reports[0];
    window.open(tokenUrl(`/api/zenkasho/report/${encodeURIComponent(first.name)}`), "_blank", "noopener");
  }
  const notes = [];
  if (v.pending.length) notes.push(`未結合の呼び出し: ${v.pending.join("・")}`);
  if (v.errors.length) notes.push(`握りつぶされたエラー(On Error Resume Next): ${v.errors.join(" / ")}`);
  $("pending").textContent = notes.join("\n");
}

function render(v) {
  V = v;
  renderTop(v);
  renderForm(v);
  renderHead(v);
  renderTable(v);
  renderRight(v);
  renderFlash(v);
  if (pack) pack.render(v.pack);
}

function remember(line) {
  recent.unshift(`${new Date().toLocaleTimeString("ja-JP")} ${line}`);
  recent.length = Math.min(recent.length, 10);
  $("recent").replaceChildren(...recent.map((l) => el("li", "", l)));
}

function init() {
  pack = packDialog({ imageUrl: "/api/zenkasho/pack_image", send: (name, body) => op(name, body) });
  $("modes").addEventListener("click", (e) => {
    const b = e.target.closest("button[data-mode]");
    if (b && b.getAttribute("aria-selected") !== "true") op("choose", { mode: b.dataset.mode });
  });
  // フォームの入力欄: 確定(change / Enter)で送る
  for (const id of ["fLot", "fKenNo", "fYouto", "fKname", "fMin", "fMax"]) {
    const i = $(id);
    i.addEventListener("change", () => op("form", { name: i.dataset.name, value: i.value }));
    i.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); i.blur(); } });
  }
  $("measIn").addEventListener("keydown", async (e) => {
    if (e.key !== "Enter" && !(e.key === "Tab" && $("measIn").value)) return;   // Enter でも Tab でも受ける
    e.preventDefault();
    const t = $("measIn").value;
    $("measIn").value = "";
    if (t === "") return;
    remember(`${V && V.active ? V.active.label : ""} ← ${t}`);
    await op("receive", { text: t });
    $("measIn").focus();
  });
  const pick = (e) => {
    const c = e.target.closest("[data-r][data-c]");
    if (c) op("select", { r: Number(c.dataset.r), c: Number(c.dataset.c) });
  };
  $("sheet").addEventListener("click", pick);
  $("head").addEventListener("click", pick);
  $("eraseBtn").addEventListener("click", () => op("erase"));
  $("emptyBtn").addEventListener("click", () => op("next-empty"));
  $("weightBtn").addEventListener("click", () => op("weight"));
  $("transferBtn").addEventListener("click", () => op("transfer"));
  $("heightBtn").addEventListener("click", () => op("height"));
  $("newBtn").addEventListener("click", () => op("new"));
  $("resumeBtn").addEventListener("click", () => op("resume"));
  $("resumeBtn2").addEventListener("click", () => op("resume"));
  // 入力欄以外を押したら、測定器の入力欄へ戻す(測定器の入力を止めない)
  document.addEventListener("click", (e) => {
    if (!e.target.closest("input, select, textarea, dialog, a")) setTimeout(focusMeas, 0);
  });
  api.get("/api/zenkasho/state").then((res) => { render(res.view); return op("start"); }).catch(toastError);
  $("measIn").focus();
}

init();

// 受信機から直接受け取る(保守で「検査画面で受け取る」にした受信機)。この画面が開いている間だけ、この画面に入る
startReceiver("zenkasho", (text) => op("receive", { text }));
