/*
  履歴画面(VBA の UF管理)。一覧と中身はサーバが作る(/api/history/*)。ここは並べるだけ。
*/

import { api } from "../api.js";
import { toast, toastError } from "../toast.js";
import * as tabs from "../tabs.js";

const $ = (id) => document.getElementById(id);
let rows = [];
let mimi = "0";
let current = 0;

function td(text, tag = "td") {
  const c = document.createElement(tag);
  c.textContent = text ?? "";
  return c;
}

function renderList() {
  const q = $("histFilter").value.trim().toUpperCase();
  const list = rows.filter((r) => !q || r.lot.toUpperCase().includes(q));
  $("histCount").textContent = `${list.length} 件`;
  $("histTable").tBodies[0].replaceChildren(...list.map((r) => {
    const tr = document.createElement("tr");
    tr.tabIndex = 0;
    tr.style.cursor = "pointer";
    tr.append(td(r.time), td(r.lot), td([r.atu, r.hab, r.tak].filter(Boolean).join(" × ")));
    tr.addEventListener("click", () => show(r.row));
    tr.addEventListener("keydown", (e) => { if (e.key === "Enter") show(r.row); });
    return tr;
  }));
}

async function load() {
  const res = await api.get(`/api/history/lots?mimi=${mimi}`);
  rows = res.rows;
  renderList();
}

async function show(row) {
  try {
    const v = await api.get(`/api/history/view?row=${row}&mimi=${mimi}`);
    for (const m of v.messages || []) toast(m.text, m.level || "info");
    $("histTitle").textContent = `検査の中身: ${v.lot}`;
    $("histNone").hidden = true;
    current = row;
    $("histRestore").hidden = false;
    const head = document.createElement("tr");
    head.append(td("", "th"), ...v.faces.map((f) => td(f, "th")));
    const wt = v.worker_time || {};
    $("histGrid").replaceChildren(head, ...v.grid.map((r) => {
      const tr = document.createElement("tr");
      tr.append(td(String(r.no), "th"), ...r.cells.map((c, i) => {
        const cell = td(c);
        const w = wt[`${v.faces[i]}${r.no}`];
        if (w && c) {                                   // 測定値の下に 検査員・時刻(2026-11-08)
          const sm = document.createElement("small");
          sm.className = "who";
          sm.textContent = [w.who, String(w.when || "").replace(/^\d{4}\/\d{1,2}\/\d{1,2}\s*/, "")].filter(Boolean).join(" ");
          cell.append(sm);
          cell.title = `測定: ${w.who || "-"}  ${w.when || ""}${w.pake ? "  パケット " + w.pake : ""}`;
        }
        return cell;
      }));
      return tr;
    }));
    $("histWho").textContent = Object.keys(wt).length
      ? "各値の下は、測った検査員と時刻です。"
      : "この検査は、測定者と時刻の控えがありません(2026-11-08 より前に印刷したもの)。";
    const head2 = document.createElement("tr");
    head2.append(td("", "th"), ...v.faces.map((f) => td(f, "th")));
    $("histExtra").replaceChildren(head2, ...Object.entries(v.extras).map(([k, list]) => {
      const tr = document.createElement("tr");
      tr.append(td(k, "th"), ...list.map((c) => td(c)));
      return tr;
    }));
    const kv = $("histOthers");
    kv.replaceChildren(...Object.entries(v.others).flatMap(([k, x]) => [td(k, "dt"), td(x, "dd")]));
  } catch (e) {
    toastError(e);
  }
}

for (const r of document.querySelectorAll('input[name="sheet"]')) {
  r.addEventListener("change", () => { mimi = r.value; load().catch(toastError); });
}
$("histFilter").addEventListener("input", renderList);
$("histRestore").addEventListener("click", async () => {
  if (!current) return;
  try {
    const r = await api.post("/api/history/restore", { row: current, mimi: mimi === "1" });
    for (const m of r.messages || []) if (m.text) toast(m.text, m.level || "info");
    if (r.stop) { toast(r.stop.text, "error"); return; }
    location.href = "/kensa";
  } catch (e) {
    toastError(e);
  }
});
load().catch(toastError);

/* ---- 理由の履歴(2026-11-10): ライン別の Access.log の「理由選択」の行を読むだけ ---- */
let rsLoaded = false;
function kvRow(k, v) { return [td(k, "dt"), td(v, "dd")]; }

async function loadReasons() {
  rsLoaded = true;
  const q = new URLSearchParams({ line: $("rsLine").value || "", from: $("rsFrom").value || "",
                                  to: $("rsTo").value || "", lot: $("rsLot").value.trim() });
  let r;
  try {
    r = await api.get(`/api/history/reasons?${q}`);
  } catch (e) {
    toastError(e);
    return;
  }
  if (!$("rsLine").options.length) {
    $("rsLine").append(...(r.lines || []).map((l) => new Option(l.label, l.key)));
  }
  $("rsLine").value = r.folder;
  $("rsFiles").replaceChildren(
    ...kvRow("理由の置き場所(ライン別 Access.log)", r.access),
    ...kvRow("測定値のログ(ライン別 Measured.log)", r.measured),
    ...kvRow("アプリの Access.log(Lot登録・測定値)", r.app_access));
  $("rsError").hidden = !r.error;
  $("rsError").textContent = r.error ? `${r.error}\n${r.access}` : "";
  $("rsError").style.whiteSpace = "pre-wrap";
  $("rsCount").textContent = r.error ? "読めません" : `${r.total} 件`;
  $("rsBadge").hidden = !!r.error || !r.total;
  $("rsBadge").textContent = r.total ? String(r.total) : "";
  $("rsTable").tBodies[0].replaceChildren(...(r.entries || []).map((e) => {
    const tr = document.createElement("tr");
    tr.dataset.kind = e.kind;
    const kind = td(e.kind); kind.className = "rs-kind";
    const why = td(e.reasons); why.className = "rs-reasons";
    tr.append(td(e.at), td(e.pc), td(e.lot || "-"), kind, td(e.position), td(e.value), why, td(e.worker));
    tr.title = e.raw;
    return tr;
  }));
  $("rsNote").textContent = [
    r.truncated ? `新しいものから ${r.entries.length} 件だけ出しています(日付で絞ってください)。` : "",
    !r.error && r.total === 0 ? "条件に合う理由はありません。" : "",
    "LotNo は、アプリの Access.log の「Lot登録」から、その時刻より前に同じ PC で登録したロットを推して出しています。",
  ].filter(Boolean).join(" ");
}

const histTabs = $("historyTabs");
tabs.attach(histTabs);
const maybeLoad = () => { if (tabs.current(histTabs) === "reasons" && !rsLoaded) loadReasons(); };
histTabs.addEventListener("click", maybeLoad);
histTabs.addEventListener("keyup", maybeLoad);
if (tabs.current(histTabs) === "reasons") loadReasons();
$("rsLoad").addEventListener("click", loadReasons);
$("rsLine").addEventListener("change", loadReasons);
$("rsLot").addEventListener("keydown", (e) => { if (e.key === "Enter") loadReasons(); });
