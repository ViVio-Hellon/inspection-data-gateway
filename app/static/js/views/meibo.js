/*
  名簿・日報の画面(VBA の UFNameList)。判断はサーバ(app/routes/meibo.py)。ここは表を描いて送るだけ。
    mode 1 検査員不一致 / mode 3 梱包枚数 / mode 2 日報閲覧(セルを直して「反映」で日報へ書き戻す)
*/
import { api } from "../api.js";
import { toast, toastError } from "../toast.js";
import { confirmBox } from "../askbox.js";

const $ = (id) => document.getElementById(id);
let V = null;
let rows = [];          // 日報閲覧で直している行(反映するまでシートは変わらない)
let selected = -1;

function el(tag, props = {}, ...kids) {
  const e = Object.assign(document.createElement(tag), props);
  e.append(...kids);
  return e;
}

async function load(mode, list = "open", open = false) {
  try {
    const res = await api.get(`/api/meibo/view?mode=${mode}&list=${list}${open ? "&open=1" : ""}`);
    draw(res.view);
  } catch (e) { toastError(e); }
}

function draw(v) {
  V = v;
  rows = v.rows.map((r) => r.slice());
  selected = -1;
  for (const b of document.querySelectorAll("#meiboTabs [data-mode]")) b.setAttribute("aria-pressed", String(Number(b.dataset.mode) === v.mode));
  $("meiboLine").textContent = [v.line, v.content].filter(Boolean).join(" ");
  $("meiboMsg").textContent = v.message || "";
  $("meiboWarn").hidden = v.mode !== 1;
  $("cb1Btn").hidden = v.mode !== 1;
  $("cb1Btn").textContent = v.list === "mismatch" ? "全ﾘｽﾄ" : "不一致ﾘｽﾄ";
  $("editBtn").hidden = v.mode !== 3;
  $("delBtn").hidden = !v.editable;
  $("postBtn").hidden = !v.editable;
  drawTable();
}

function drawTable() {
  const v = V;
  const head = el("tr", {}, ...v.head.map((h) => el("th", { textContent: h })));
  const body = rows.map((r, i) => {
    const tr = el("tr");
    if (i === selected) tr.className = "is-selected";
    tr.addEventListener("click", () => { selected = i; for (const x of $("meiboTable").querySelectorAll("tr.is-selected")) x.className = ""; tr.className = "is-selected"; });
    r.forEach((c, j) => {
      if (v.editable) {
        const inp = el("input", { className: "input", value: c, ariaLabel: v.head[j] || "" });
        inp.style.minWidth = "5em";
        inp.addEventListener("input", () => { rows[i][j] = inp.value; });
        tr.append(el("td", {}, inp));
      } else {
        tr.append(el("td", { textContent: c }));
      }
    });
    return tr;
  });
  $("meiboTable").replaceChildren(el("thead", {}, head), el("tbody", {}, ...body));
}

async function post(answers = {}) {
  let res;
  try { res = await api.post("/api/meibo/daily", { rows, answers }); } catch (e) { toastError(e); return; }
  if (res.question) {
    const yes = await confirmBox(res.question.text, { ok: "はい", cancel: "いいえ" });
    return post({ ...answers, [res.question.key]: !!yes });
  }
  if (res.stop) { toast(res.stop.text, "error"); return; }
  draw(res.view);
  if (answers.daily) toast("日報を直しました", "info");
}

function init() {
  for (const b of document.querySelectorAll("#meiboTabs [data-mode]")) b.addEventListener("click", () => load(Number(b.dataset.mode)));
  $("cb1Btn").addEventListener("click", () => load(1, V && V.list === "mismatch" ? "all" : "mismatch"));
  $("editBtn").addEventListener("click", () => load(2, "open", true));     // OnOff「編集起動」→ NameList_Open(2)
  $("delBtn").addEventListener("click", async () => {
    if (selected < 0) return;
    // UFNameList.CommandButton1: 選んだ行の .SelectedItem(先頭の空の列)を出すので名前は空になる(VBA どおり)
    if (await confirmBox("現在選択されている  を削除しますか？", { ok: "はい", cancel: "いいえ" })) {
      rows.splice(selected, 1);
      selected = -1;
      drawTable();
    }
  });
  $("postBtn").addEventListener("click", () => post());
  const mode = Number(new URLSearchParams(location.search).get("mode") || "3");
  load([1, 2, 3].includes(mode) ? mode : 3, "open", true);
}

init();
