/*
  取り込み画面。一覧はサーバが作る(`/api/import/plan`)。ここは並べて、選んだものを送るだけ。
  進み具合は帯の「いま動いているもの」(jobs.js)に出る。
*/

import { api } from "../api.js";
import { toast, toastError } from "../toast.js";

const $ = (id) => document.getElementById(id);
let P = null;

function td(text) {
  const c = document.createElement("td");
  c.textContent = text;
  return c;
}

function when(t) {
  if (!t) return "まだ";
  const d = new Date(t * 1000);
  return d.toLocaleString("ja-JP");
}

function render() {
  const body = $("importTable").tBodies[0];
  body.replaceChildren(...P.rows.map((r) => {
    const tr = document.createElement("tr");
    const pick = document.createElement("td");
    const box = document.createElement("input");
    box.type = "checkbox";
    box.value = r.name;
    box.disabled = !P.admin || !P.can_run;
    box.setAttribute("aria-label", r.name + " を選ぶ");
    pick.append(box);
    const src = document.createElement("td");
    const input = document.createElement("input");
    input.className = "input";
    input.style.minWidth = "28em";
    input.value = r.source;
    input.dataset.name = r.name;
    input.disabled = !P.admin;
    src.append(input, document.createElement("br"));
    const note = document.createElement("small");
    note.className = r.source_exists ? "muted" : "lv--ng";
    note.textContent = r.source_exists ? "あります" : "見つかりません";
    src.append(note);
    tr.append(pick, td(r.name), src, td(r.target), td(when(r.target_mtime)));
    return tr;
  }));
  $("importState").textContent = P.can_run ? (P.admin ? "取り込めます" : "管理者だけ") : "この PC では取り込めません";
  $("importState").className = "lv " + (P.can_run && P.admin ? "lv--ok" : "lv--info");
  $("importWhy").textContent = P.why || (P.admin ? "" : "取り込むには管理者(ufkensa:admin)の権限が要ります。");
  $("importRun").disabled = !P.admin || !P.can_run;
  $("importAll").disabled = !P.admin || !P.can_run;
  $("xlsmRun").disabled = !P.admin;
}

async function load() {
  P = await api.get("/api/import/plan");
  render();
}

async function runImport() {
  const rows = [...$("importTable").tBodies[0].rows];
  const names = rows.map((tr) => tr.querySelector("input[type=checkbox]")).filter((b) => b.checked).map((b) => b.value);
  const sources = {};
  for (const input of $("importTable").querySelectorAll("input.input")) sources[input.dataset.name] = input.value;
  try {
    await api.post("/api/import/run", { names, sources });
    toast("取り込みを始めました。進み具合は上の帯に出ます。", "info");
  } catch (e) {
    toastError(e);
  }
}

$("importAll").addEventListener("change", () => {
  for (const b of $("importTable").querySelectorAll("tbody input[type=checkbox]")) b.checked = $("importAll").checked;
});
$("importRun").addEventListener("click", runImport);
$("xlsmRun").addEventListener("click", async () => {
  try {
    await api.post("/api/import/xlsm", { path: $("xlsmPath").value });
    toast("xlsm の取り込みを始めました。進み具合は上の帯に出ます。", "info");
  } catch (e) {
    toastError(e);
  }
});
load().catch(toastError);
