/*
  キャスク検査(VBA の UFCask + キャスク検査シート)。
  判断はすべてサーバ(ufkensa/kensa/cask_session.py → modes/cask.py・sheet.py)。ここは受け取った view を描き、操作を送るだけ。

  操作の返り値(Result): { view, messages, question, stop }(検査画面と同じ形)
  表は 700 行のうち入力のある所 +α だけ届く。行ごとに中身が変わった所だけ描き直し、次に入力する行が見えるようにスクロールする。
*/

import { mayRefocus } from "../focus_guard.js";
import { sheetInput } from "../sheet_input.js";
import { openWindow } from "../desktop.js";
import { api, tokenUrl } from "../api.js";
import { toastError } from "../toast.js";
import { screenOp } from "../screen_op.js";
import { packDialog } from "./pack_dialog.js";
import { startReceiver } from "../recv.js";
import { canSpeak, readOn, setReadOn, speak, voiceName } from "../speak.js";

const $ = (id) => document.getElementById(id);
let sheetIn = null;
const recent = [];
let V = null;
let colSig = "";
const rowEls = new Map();      // 行番号 → { tr, sig }
let lastReport = "";
let pack = null;               // 包装仕様書の窓(UF_PackNo)

/* 操作を送る(screen_op.js)。終わったら測定の欄へ戻す({ refocus: false } で戻さない) */
const op = screenOp("cask", {
  render,
  after: (name, res, { refocus = true } = {}) => { if (refocus) focusInput(); },
});

function focusInput() {
  if (!V || !V.visible) return;
  if (document.querySelector("dialog[open]") || !mayRefocus()) return;
  // 2026-11-26: ゼロ校正も同じ欄で受ける(サーバがゼロ校正へ回す)
  $("measIn").focus({ preventScroll: true });
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
  // 2026-11-26(現場の声「色が薄くて ON/OFF が分かりにくい」): ON は濃い緑の大きな札とボタン、OFF は灰色
  const on = f.inspection === "ON";
  text($("inspLv"), on ? "ON(巾・丈・対角を測る)" : "OFF(厚・外観だけ)");
  $("inspLv").className = "insp-state " + (on ? "insp-state--on" : "insp-state--off");
  $("inspBtn").textContent = on ? "巾丈測定を OFF にする" : "巾丈測定を ON にする";
  $("inspBtn").setAttribute("aria-pressed", String(on));
  const zeroOk = f.micro !== "" && f.convex !== "";
  text($("zeroLv"), zeroOk ? "校正済" : "未");
  $("zeroLv").className = "lv " + (zeroOk ? "lv--ok" : "lv--ng");
  // 2026-11-24: ゼロ校正がまだなら画面の上に大きく(測ると、測定ではなくゼロ校正に入る)
  const need = [[f.l1 || "マイクロ/ノギス", f.micro], ["コンベックス", f.convex]].filter(([, v]) => v === "").map(([n]) => n);
  $("zeroAlert").hidden = zeroOk || !V.visible;
  text($("zeroAlertNeed"), need.length ? `まだ: ${need.join("・")}。いま測ると、ゼロ校正に入ります(測定値は表に入りません)` : "");
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
    lastFlash = v.flash;
    $("flash").dataset.mark = v.flash.mark;
    text($("flashMark"), v.flash.mark);
    text($("flashText"), [v.flash.no && `No.${v.flash.no}`, v.flash.label].filter(Boolean).join(" "));
  }
  $("labels").replaceChildren(...v.form.labels.map((l) => { const li = document.createElement("li"); li.textContent = l; return li; }));
  kv($("mmKv"), v.header.maxmin.map((m) => [m.name, m.max || m.min ? `MAX ${m.max} / MIN ${m.min}` : ""]));
  $("ainLink").hidden = !v.ainput;
  $("ainLink").href = "/ainput?from=Cask&method=L2";            // 呼び元のシートのロット・Min/Max で開く
  const rep = v.report;
  $("reportLink").hidden = !rep;
  if (rep) {
    const url = tokenUrl(`/api/cask/report/${encodeURIComponent(rep.name)}`);
    $("reportLink").href = url;
    if (lastReport && lastReport !== rep.html + rep.name) {
      openWindow(url, "キャスク検査");
      // 2026-11-15: 不良要因票も一緒に開く(通常検査と同じく 2 枚刷る)
      if (rep.defect) openWindow(tokenUrl(`/api/cask/report/${encodeURIComponent(rep.defect)}`), "不良要因票");
    }
  }
  lastReport = rep ? rep.html + rep.name : "none";
  text($("savedNote"), v.saved ? `${v.saved.name}: ${v.saved.files.length} ファイル` : "");
  $("closed").hidden = v.visible;
  text($("closedText"), "キャスク検査は開いていません。検査画面で検査員を登録してから「開く」を押してください。");
}

/* ---- 読み上げ(UFCask.Read / CommandButton14。2026-11-08) ---- */
let lastFlash = null;
function flashWords(f) { return f ? [f.no, f.label].filter(Boolean).join("、") : ""; }
async function receive(t) {
  const r = await op("receive", { text: t });
  const f = r && r.view && r.view.flash;
  if (f) { lastFlash = f; if (readOn()) say(flashWords(f)); }
}
/* 2026-11-26(現場の声「読み上げ ON/OFF が効いているか分からない」):
   ON にしたら「読み上げ オン」と声で返す・使う声の名前と読んでいる最中を出す・読めない PC はそう出す */
let readState = "";
function renderRead() {
  const on = readOn();
  const b = $("readToggle");
  b.setAttribute("aria-pressed", String(on));
  b.textContent = on ? "読み上げ ON" : "読み上げ OFF";
  b.classList.toggle("btn--commit", on);
  b.disabled = $("readNow").disabled = !canSpeak();
  let note;
  if (!canSpeak()) note = "この画面では読み上げできません(読み上げの機能がありません)";
  else if (readState) note = readState;
  else if (!on) note = "OFF: 測っても読み上げません";
  else note = voiceName() ? `ON: 測るたびに読み上げます(声 ${voiceName()})` : "ON: 日本語の声が見つかりません。Windows の設定 → 時刻と言語 → 音声 で日本語の声を入れてください";
  text($("readNote"), note);
}
function say(words) {
  if (!words) return;
  speak(words, {
    onstart: () => { readState = `読み上げ中: ${words}`; renderRead(); },
    onend: (e) => { readState = e && e.error && e.error !== "interrupted" && e.error !== "canceled" ? `読み上げできませんでした(${e.error})` : ""; renderRead(); },
  });
}
$("readToggle").addEventListener("click", () => {
  const on = !readOn();
  setReadOn(on);
  if (on) say("読み上げ オン");
  else { try { window.speechSynthesis.cancel(); } catch { /* 読めない画面 */ } readState = ""; }
  renderRead();
});
$("readNow").addEventListener("click", () => say(lastFlash ? flashWords(lastFlash) : "まだ測っていません"));
if ("speechSynthesis" in window) window.speechSynthesis.addEventListener?.("voiceschanged", renderRead);
renderRead();

function render(v) {
  V = v;
  queueMicrotask(() => sheetIn && sheetIn.refresh());
  renderTop(v);
  renderGrid(v);
  renderSide(v);
  if (pack) pack.render(v.pack);
  renderDefect(v);
}

/* ---------- 不良要因(2026-11-15) ---------- */
let defSel = -1;
function renderDefect(v) {
  const d = v.defect || { items: [], total: 0 };
  $("defLv").textContent = d.total ? `${d.total}枚` : "";
  $("defLot").textContent = d.lot ? `Lot ${d.lot}` : "ロット未登録";
  $("defWarn").hidden = !d.other_lot;
  $("defWarn").textContent = d.other_lot
    ? `いまの不良要因票は別のロット(${d.other_lot})のものです。反映すると、このロットの票を作り直します。` : "";
  $("defList").replaceChildren(...d.items.map((it) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "cdef__item";
    b.setAttribute("role", "option");
    b.setAttribute("aria-selected", String(it.i === defSel));
    const n = document.createElement("span"); n.textContent = it.name;
    const c = document.createElement("b"); c.textContent = it.count ? it.count : "";
    b.append(n, c);
    b.addEventListener("click", () => {
      defSel = it.i;
      for (const x of $("defList").children) x.setAttribute("aria-selected", String(x === b));
      $("defCount").focus();
    });
    return b;
  }));
}
function defPreview() {
  $("defPrev").src = tokenUrl(`/api/cask/defect_preview?_=${Date.now()}`);
  $("defOpen").href = tokenUrl("/api/cask/defect_preview");
}
async function defReflect() {
  const count = $("defCount").value.trim();
  await op("defect-reflect", { item: defSel, count }, undefined, { refocus: false });
  $("defCount").value = "";
  defPreview();
  $("defCount").focus();
}

function remember(line) {
  recent.unshift(`${new Date().toLocaleTimeString("ja-JP")} ${line}`);
  recent.length = Math.min(recent.length, 8);
  $("recent").replaceChildren(...recent.map((l) => { const li = document.createElement("li"); li.textContent = l; return li; }));
}

/* 受け取れるか(測定器の文字はカーソルのある所へ入る)を色で出す。通常検査と同じ */
function recvShow() {
  const a = document.activeElement;
  const on = a === $("measIn");
  $("recvBox").dataset.on = on ? "1" : "";
  text($("recvState"), on ? "受信待ち(選んだセルに入ります)" : "ほかの欄に入力中 — 表のセルを押すと測定を受けます");
}

/* ---------- 外観(2026-11-26。押すと窓で入れる) ---------- */
const APPEARANCE_COL = 28;
let appRow = 0;
function openAppearance(row, current) {
  appRow = row;
  const no = row - 6;                                       // A 列の板No(7 行目が 1)
  $("appTitle").textContent = `外観(板No ${no})`;
  $("appIn").value = current;
  $("appDlg").showModal();
  $("appIn").focus();
  $("appIn").select();
}
async function commitAppearance(text) {
  $("appDlg").close();
  await op("select", { row: appRow, col: APPEARANCE_COL }, undefined, { refocus: false });
  if (text === "") await op("clear");
  else { remember(`外観 ← ${text}`); await op("receive", { text }); }
}
function initAppearance() {
  $("appOk").addEventListener("click", () => commitAppearance($("appIn").value.trim()));
  $("appClear").addEventListener("click", () => commitAppearance(""));
  $("appCancel").addEventListener("click", () => { $("appDlg").close(); focusInput(); });
  $("appIn").addEventListener("keydown", (e) => {
    if (e.isComposing || e.keyCode === 229) return;          // 日本語の変換の確定では送らない
    if (e.key === "Enter") { e.preventDefault(); commitAppearance($("appIn").value.trim()); }
  });
}

function initDefect() {
  $("defBtn").addEventListener("click", () => { defPreview(); $("defDlg").showModal(); });
  $("defClose").addEventListener("click", () => { $("defDlg").close(); focusInput(); });
  $("defReflect").addEventListener("click", defReflect);
  $("defCount").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); defReflect(); } });
  $("defClear").addEventListener("click", async () => {
    await op("defect-clear", {}, undefined, { refocus: false });
    defPreview();
  });
}

function init() {
  initDefect();
  initAppearance();
  pack = packDialog({ imageUrl: "/api/cask/pack_image", send: (name, body) => op(name, body) });
  $("measIn").addEventListener("focus", recvShow);
  $("measIn").addEventListener("blur", () => setTimeout(recvShow, 0));
  // 2026-11-26: 表は Excel の良いところだけ(選んだセルに打つ・矢印・Delete・Back space)。sheet_input.js
  sheetIn = sheetInput({
    input: $("measIn"),
    cell: () => document.querySelector('#caskGrid td[data-sel="1"]') || document.querySelector('#tolTable td[data-sel="1"]'),
    send: (t) => { remember(t); return receive(t); },
    backspace: () => op("cancel"),
    del: () => op("clear"),
    move: (dr, dc) => {
      if (!V || !V.next) return null;
      const row = V.next.row + dr, col = V.next.col + dc;
      if (row < 1 || col < 2) return null;
      return op("select", { row, col });
    },
  });
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
    const col = Number(td.dataset.col);
    if (!row) return;
    // 2026-11-26(現場の声): 外観は押すと別の窓で入れる(表の上に打つより入れやすい)
    if (col === APPEARANCE_COL && row >= 7) { openAppearance(row, td.textContent.trim()); return; }
    op("select", { row, col });
  };
  $("caskGrid").addEventListener("click", pick);
  $("tolTable").addEventListener("click", pick);
  $("inspBtn").addEventListener("click", () => op("inspection"));
  $("proofBtn").addEventListener("click", () => op("proofread"));
  $("transferBtn").addEventListener("click", () => op("transfer"));
  $("printBtn").addEventListener("click", () => op("print"));
  $("saveBtn").addEventListener("click", () => op("save"));
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

// 受信機から直接受け取る(保守で「検査画面で受け取る」にした受信機)。この画面が開いている間だけ、この画面に入る
startReceiver("cask", (text) => receive(text));
