/*
  キャスク検査(VBA の UFCask + キャスク検査シート)。
  判断はすべてサーバ(ufkensa/kensa/cask_session.py → modes/cask.py・sheet.py)。ここは受け取った view を描き、操作を送るだけ。

  操作の返り値(Result): { view, messages, question, stop }(検査画面と同じ形)
  表は 700 行のうち入力のある所 +α だけ届く。行ごとに中身が変わった所だけ描き直し、次に入力する行が見えるようにスクロールする。
*/

import { api, tokenUrl } from "../api.js";
import { toast, toastError } from "../toast.js";
import { playMessage } from "../sound.js";
import { confirmBox } from "../askbox.js";
import { packDialog } from "./pack_dialog.js";

const $ = (id) => document.getElementById(id);
const LEVEL_KIND = { info: "info", warn: "warn", error: "error" };
const recent = [];
let V = null;
let colSig = "";
const rowEls = new Map();      // 行番号 → { tr, sig }
let lastReport = "";
let pack = null;               // 包装仕様書の窓(UF_PackNo)

async function op(name, body = {}, answers = {}, { refocus = true } = {}) {
  let res;
  try {
    res = await api.post(`/api/cask/${name}`, { ...body, answers });
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
    const yes = await confirmBox(q.text, { ok: q.buttons === "okcancel" ? "OK" : "はい",
                                           cancel: q.buttons === "okcancel" ? "キャンセル" : "いいえ" });
    return op(name, body, { ...answers, [q.key]: !!yes }, { refocus });
  }
  if (res.stop) toast(res.stop.text, "error");
  if (refocus) focusInput();
  return res;
}

function focusInput() {
  if (!V || !V.visible) return;
  if (document.querySelector("dialog[open]")) return;
  const el = V.focus === "zero" ? $("zeroIn") : $("measIn");
  el.focus();
}

function text(el, t) { el.textContent = t ?? ""; }

function kv(dl, pairs) {
  dl.replaceChildren(...pairs.flatMap(([k, v]) => {
    const dt = document.createElement("dt"); dt.textContent = k;
    const dd = document.createElement("dd"); dd.textContent = v ?? "";
    return [dt, dd];
  }));
}

function setField(id, value) {
  const el = $(id);
  if (document.activeElement !== el) el.value = value ?? "";
}

// ---- 上 ---------------------------------------------------------------
function renderTop(v) {
  const f = v.form, h = v.header;
  text($("lineLv"), f.line_name);
  setField("lotIn", f.lot_box);
  for (const k of ["ken_no", "kname", "cno", "mno", "seri_no", "tai", "min_text", "max_text"]) setField(`f_${k}`, f[k]);
  // 決まったロット(D2)を大きく。入力欄は控えめ
  $("lotReg").hidden = !h.lot;
  $("lotEmpty").hidden = !!h.lot;
  $("lotRegNo").textContent = h.lot || "";
  kv($("lotRegKv"), [["検査番号", f.ken_no], ["開始", h.start]]);
  kv($("lotKv"), []);

  const sel = v.next ? `${v.next.row},${v.next.col}` : "";
  const t = $("tolTable");
  const head = document.createElement("tr");
  head.append(...["", "寸法", "−", "+", "下限 〜 上限"].map((s) => { const th = document.createElement("th"); th.textContent = s; return th; }));
  const rows = h.tolerance.map((x) => {
    const tr = document.createElement("tr");
    const th = document.createElement("th"); th.textContent = x.name;
    const cell = (val, rc) => {
      const td = document.createElement("td");
      td.textContent = val;
      if (rc) { td.dataset.row = rc[0]; td.dataset.col = rc[1]; td.dataset.sel = sel === rc.join(",") ? "1" : ""; }
      return td;
    };
    tr.append(th, cell(x.size, x.size_cell), cell(x.minus, x.minus_cell), cell(x.plus, x.plus_cell),
              cell(x.lo || x.hi ? `${x.lo} 〜 ${x.hi}` : ""));
    return tr;
  });
  t.replaceChildren(head, ...rows);

  text($("countSt"), f.count_st || "-");
  text($("inspLv"), f.inspection);
  $("inspLv").className = "lv " + (f.inspection === "ON" ? "lv--ok" : "");
  const zeroOk = f.micro !== "" && f.convex !== "";
  text($("zeroLv"), zeroOk ? "校正済" : "未");
  $("zeroLv").className = "lv " + (zeroOk ? "lv--ok" : "lv--ng");
  kv($("zeroKv"), [[f.l1 || "ﾏｲｸﾛｾﾞﾛ点", f.micro || "未"], ["ｺﾝﾍﾞｯｸｽｾﾞﾛ点", f.convex || "未"],
                   ["ｺﾝﾍﾞｯｸｽ1000mm", f.convex_1000 || "未"]]);
}

// ---- 表 ---------------------------------------------------------------
function buildHead(cols) {
  const t = $("caskGrid");
  const groups = document.createElement("tr");
  const labels = document.createElement("tr");
  const g0 = document.createElement("th"); g0.className = "no"; g0.textContent = "";
  const l0 = document.createElement("th"); l0.className = "no"; l0.textContent = "No";
  groups.append(g0); labels.append(l0);
  let prev = null, span = null;
  for (const c of cols) {
    const n = (c.judge ? 2 : 1);
    if (c.hidden) continue;
    if (c.group !== prev) {
      span = document.createElement("th"); span.textContent = c.group; span.colSpan = n;
      groups.append(span); prev = c.group;
    } else { span.colSpan += n; }
    const th = document.createElement("th"); th.textContent = c.label; th.colSpan = n;
    labels.append(th);
  }
  const thead = document.createElement("thead");
  thead.append(groups, labels);
  const tbody = document.createElement("tbody");
  tbody.id = "caskBody";
  t.replaceChildren(thead, tbody);
  rowEls.clear();
}

function buildRow(r, cols) {
  const [row, no, cells] = r;
  const tr = document.createElement("tr");
  tr.dataset.row = row;
  const th = document.createElement("th"); th.className = "no"; th.textContent = no;
  tr.append(th);
  cells.forEach(([val, mark, flags], i) => {
    const c = cols[i];
    if (c.hidden) return;
    const td = document.createElement("td");
    td.dataset.col = c.col;
    td.className = flags + (c.col === 21 ? " mark" : "") + (c.col === 23 ? " note" : "");
    td.textContent = val;
    tr.append(td);
    if (c.judge) {
      const tj = document.createElement("td");
      tj.dataset.col = c.judge;
      tj.className = "mark " + flags.split(" ").filter((x) => x === "ok" || x === "ng").join(" ");
      tj.textContent = mark;
      tr.append(tj);
    }
    if (flags.includes("sel")) tr.classList.add("cur");
  });
  return tr;
}

function renderGrid(v) {
  const sig = JSON.stringify(v.columns);
  if (sig !== colSig) { buildHead(v.columns); colSig = sig; }
  const body = $("caskBody");
  const seen = new Set();
  for (const r of v.rows) {
    const key = r[0];
    seen.add(key);
    const rs = JSON.stringify(r);
    const old = rowEls.get(key);
    if (old && old.sig === rs) continue;
    const tr = buildRow(r, v.columns);
    if (old) old.tr.replaceWith(tr); else body.append(tr);
    rowEls.set(key, { tr, sig: rs });
  }
  for (const [key, o] of rowEls) {
    if (!seen.has(key)) { o.tr.remove(); rowEls.delete(key); }
  }
  let more = $("moreRow");
  if (!more) { more = document.createElement("tr"); more.id = "moreRow"; more.className = "more"; more.append(document.createElement("td")); }
  more.firstChild.colSpan = 40;
  const last = v.rows.length ? v.rows[v.rows.length - 1][0] - 6 : 0;
  more.firstChild.textContent = last < v.total_rows ? `… ${v.total_rows} 枚まで(入力に合わせて表示します)` : "";
  body.append(more);
  text($("rowsLv"), `${v.rows.length} / ${v.total_rows} 行を表示`);
  scrollToNext(v);
}

/** 次に入力する行が見えていなければ、表の中ほどに来るようにスクロールする。 */
function scrollToNext(v) {
  if (!v.next) return;
  const o = rowEls.get(v.next.row);
  if (!o) return;
  const box = $("scroll");
  const head = $("caskGrid").tHead ? $("caskGrid").tHead.offsetHeight : 0;
  const top = o.tr.offsetTop, h = o.tr.offsetHeight;
  if (top - head < box.scrollTop + h || top + h * 2 > box.scrollTop + box.clientHeight) {
    box.scrollTop = Math.max(0, top - box.clientHeight / 2);
  }
}

// ---- 右 ---------------------------------------------------------------
function renderSide(v) {
  const n = v.next;
  text($("nextName"), n ? (n.no ? `No.${n.no}  ${n.label}` : n.label) : "-");
  text($("nextHint"), n && n.locked ? "このセルは保護されています(入力できません)" : "");
  $("flash").hidden = !v.flash;
  if (v.flash) {
    $("flash").dataset.mark = v.flash.mark;
    text($("flashMark"), v.flash.mark);
    text($("flashText"), [v.flash.no && `No.${v.flash.no}`, v.flash.label].filter(Boolean).join(" "));
  }
  $("eventsBanner").hidden = v.events || !v.visible;
  $("labels").replaceChildren(...v.form.labels.map((l) => { const li = document.createElement("li"); li.textContent = l; return li; }));
  kv($("mmKv"), v.header.maxmin.map((m) => [m.name, m.max || m.min ? `MAX ${m.max} / MIN ${m.min}` : ""]));
  $("ainLink").hidden = !v.ainput;
  const rep = v.report;
  $("reportLink").hidden = !rep;
  if (rep) {
    const url = tokenUrl(`/api/cask/report/${encodeURIComponent(rep.name)}`);
    $("reportLink").href = url;
    if (lastReport && lastReport !== rep.html + rep.name) window.open(url, "_blank", "noopener");
  }
  lastReport = rep ? rep.html + rep.name : "none";
  text($("savedNote"), v.saved ? `${v.saved.name}: ${v.saved.files.length} ファイル` : "");
  $("closed").hidden = v.visible;
  text($("closedText"), "キャスク検査は開いていません。検査画面で検査員を登録してから「開く」を押してください。");
}

function render(v) {
  V = v;
  renderTop(v);
  renderGrid(v);
  renderSide(v);
  if (pack) pack.render(v.pack);
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
  pack = packDialog({ imageUrl: "/api/cask/pack_image", send: (name, body) => op(name, body) });
  onEnter("measIn", (t) => { remember(t); op("receive", { text: t }); });
  onEnter("zeroIn", (t) => op("zero", { text: t }));
  $("lotIn").addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    op("lot", { lot: $("lotIn").value });
  });
  // VBA の UFCask.LotBox_Change と同じく、7 桁まで打ったら Enter を待たずに読む
  $("lotIn").addEventListener("input", () => { if ($("lotIn").value.trim().length === 7) op("lot", { lot: $("lotIn").value }); });
  for (const el of document.querySelectorAll("[data-field]")) {
    el.addEventListener("change", () => op("field", { name: el.dataset.field, value: el.value }, {}, { refocus: false }));
    el.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); el.blur(); focusInput(); } });
  }
  const pick = (e) => {
    const td = e.target.closest("td[data-col]");
    const tr = e.target.closest("[data-row]");
    if (!td) return;
    const row = Number(td.dataset.row || (tr && tr.dataset.row));
    if (row) op("select", { row, col: Number(td.dataset.col) });
  };
  $("caskGrid").addEventListener("click", pick);
  $("tolTable").addEventListener("click", pick);
  $("inspBtn").addEventListener("click", () => op("inspection"));
  $("proofBtn").addEventListener("click", () => op("proofread"));
  $("resumeBtn").addEventListener("click", () => op("resume"));
  $("transferBtn").addEventListener("click", () => op("transfer"));
  $("printBtn").addEventListener("click", () => op("print"));
  $("saveBtn").addEventListener("click", () => op("save"));
  $("backBtn").addEventListener("click", () => op("cancel"));
  $("newBtn").addEventListener("click", () => op("new"));
  $("openBtn").addEventListener("click", () => op("open"));
  $("closeBtn").addEventListener("click", async () => { await op("close", {}, {}, { refocus: false }); location.href = "/kensa"; });
  api.get("/api/cask/state").then((res) => {
    render(res.view);
    if (!res.view.visible) return op("open");
    focusInput();
    return null;
  }).catch(toastError);
}

init();
