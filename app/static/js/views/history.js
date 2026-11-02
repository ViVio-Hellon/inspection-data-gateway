/*
  履歴画面(VBA の UF管理)。一覧と中身はサーバが作る(/api/history/*)。ここは並べるだけ。
*/

import { api } from "../api.js";
import { toast, toastError } from "../toast.js";

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
    $("histGrid").replaceChildren(head, ...v.grid.map((r) => {
      const tr = document.createElement("tr");
      tr.append(td(String(r.no), "th"), ...r.cells.map((c) => td(c)));
      return tr;
    }));
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
