/*
  保守画面(管理者だけ)。歪公差チェック・ソリ公差・設備仕様・端末設定のセット。
  値の書式・状態・違いの有無はサーバ(ufkensa/kensa/maint.py)が決める。ここは描いて送るだけ。
*/

import { api } from "../api.js";
import { confirmBox } from "../askbox.js";
import { toast, toastError } from "../toast.js";

const $ = (id) => document.getElementById(id);
const STATUS = { ok: ["", ""], none: ["該当なし", "lv--warn"], error: ["エラー", "lv--ng"], not_found: ["ロットなし", "lv--ng"] };
// 一覧に出す欄(コメントは長いので最後・折り返し)
const LIST_KEYS = ["MG", "HaHiz", "TaHiz", "Naka", "Mimi", "NoTex", "Source", "Flag", "HizCome"];

let S = { labels: {}, keys: [] };
let lastAll = null;

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text ?? "";
  return e;
}

function kv(dl, pairs) {
  dl.replaceChildren(...pairs.flatMap(([k, v, cls]) => {
    const dd = el("dd", cls || "", v === "" || v == null ? "-" : v);
    return [el("dt", "", k), dd];
  }));
}

function statusChip(st) {
  const [label, cls] = STATUS[st] || [st, ""];
  return label ? el("span", "lv " + cls, label) : el("span", "lv lv--ok", "OK");
}

function altText() { return $("mAlt").value.trim(); }

/* ---------- 歪公差: 1 ロット(縦に項目、横に いまの表 / 比べる表) ---------- */
function renderOne(d) {
  const r = d.row;
  const cur = r.current, alt = r.alt;
  $("mSummary").replaceChildren(
    el("b", "", r.lot), el("span", "muted", "検番"), el("code", "", r.ken_no || "-"),
    el("span", "muted", "引当"), el("span", "", (r.hiki || []).map((h) => h[0]).join(" / ") || "-"),
    statusChip(cur.status),
    ...(alt ? [el("span", "lv " + (r.diff.length ? "lv--warn" : "lv--ok"), r.diff.length ? `違い ${r.diff.length} 項目` : "違いなし")] : []),
  );
  const t = el("table", "table maint__one");
  const head = el("tr");
  head.append(el("th", "", "項目"), el("th", "", "いまの品質基準"));
  if (alt) head.append(el("th", "", "比べる品質基準"));
  const thead = el("thead"); thead.append(head);
  const tbody = el("tbody");
  for (const k of S.keys) {
    const tr = el("tr");
    if (r.diff.includes(k)) tr.className = "is-diff";
    tr.append(el("th", "", S.labels[k] + (k === "Flag" ? "(参考)" : "")), el("td", "", cur.values[k] ?? ""));
    if (alt) tr.append(el("td", "", alt.values[k] ?? ""));
    tbody.append(tr);
  }
  t.append(thead, tbody);
  const box = [t];
  for (const [name, x] of [["いまの品質基準", cur], ["比べる品質基準", alt]]) {
    if (!x) continue;
    if (x.note) box.push(el("p", "maint__note", `${name}: ${x.note}`));
    for (const m of x.messages || []) box.push(el("p", "maint__msg", `${name}: ${m}`));
  }
  for (const m of r.messages || []) box.push(el("p", "maint__msg", m));
  $("mResult").replaceChildren(...box);
}

/* ---------- 歪公差: 全ロット(1 行 1 ロット。違う欄は「いま → 比べる」) ---------- */
function cell(r, k) {
  const td = el("td", ["HizCome", "Source", "Flag"].includes(k) ? "maint__wrap" : "n");
  const a = r.current.values[k] ?? "";
  if (r.alt && r.diff.includes(k)) {
    td.classList.add("is-diff");
    td.append(el("span", "", a || "(空)"), el("span", "maint__arrow", " → "), el("b", "", (r.alt.values[k] ?? "") || "(空)"));
  } else {
    td.textContent = a;
  }
  return td;
}

function renderAll(d) {
  lastAll = d;
  const b = d.by_status;
  $("mSummary").replaceChildren(
    el("b", "", `${d.count} ロット`), el("span", "muted", `${d.sec} 秒`),
    el("span", "lv lv--ok", `OK ${b.ok}`), el("span", "lv lv--warn", `該当なし ${b.none}`),
    el("span", "lv lv--ng", `エラー ${b.error + b.not_found}`),
    ...(d.alt ? [el("span", "lv " + (d.diffs ? "lv--warn" : "lv--ok"), `違いのあるロット ${d.diffs}`)] : []),
  );
  const onlyDiff = $("mOnlyDiff").checked;
  const t = el("table", "table table--dense maint__all");
  const head = el("tr");
  head.append(el("th", "", "ロット"), el("th", "", "検番"), el("th", "", "状態"), ...LIST_KEYS.map((k) => el("th", "", S.labels[k])));
  const thead = el("thead"); thead.append(head);
  const tbody = el("tbody");
  for (const r of d.rows) {
    if (onlyDiff && !r.diff.length) continue;
    const tr = el("tr");
    if (r.diff.length) tr.className = "is-diffrow";
    const st = el("td");
    st.append(statusChip(r.current.status));
    if (r.alt && r.diff.includes("status")) st.append(el("span", "maint__arrow", " → "), statusChip(r.alt.status));
    if (r.current.note) st.title = r.current.note;
    tr.append(el("td", "n", r.lot), el("td", "n", r.ken_no), st, ...LIST_KEYS.map((k) => cell(r, k)));
    tbody.append(tr);
  }
  t.append(thead, tbody);
  $("mResult").replaceChildren(t);
}

$("mOne").addEventListener("submit", async (e) => {
  e.preventDefault();
  const lot = $("mLot").value.trim().toUpperCase();
  $("mLot").value = lot;
  try {
    renderOne(await api.post("/api/maint/dist", { lot, alt: altText() }));
    toast("ロード済", "ok");                             // Test:107 MsgBox "ロード済"
  } catch (err) { toastError(err); }
});

$("mAll").addEventListener("submit", async (e) => {
  e.preventDefault();
  try {
    renderAll(await api.post("/api/maint/dist_all", { prefix: $("mPrefix").value.trim(), plate_only: $("mPlate").checked, alt: altText() }));
  } catch (err) { toastError(err); }
});
$("mOnlyDiff").addEventListener("change", () => { if (lastAll) renderAll(lastAll); });

/* ---------- 比べる品質基準: 参照(/api/fs/list) ---------- */
async function browse(path) {
  try {
    const d = await api.get(`/api/fs/list?path=${encodeURIComponent(path || "")}`);
    $("mBPath").textContent = d.path || "(場所を選ぶ)";
    const items = [];
    const li = (text, cls, fn) => { const x = el("li", cls); const b = el("button", "", text); b.type = "button"; b.addEventListener("click", fn); x.append(b); return x; };
    if (d.parent) items.push(li("↑ 上へ", "is-dir", () => browse(d.parent)));
    for (const r of d.path ? [] : d.roots) items.push(li(r.name || r.path, "is-dir", () => browse(r.path)));
    for (const x of d.dirs) items.push(li(x.name, "is-dir", () => browse(x.path)));
    for (const x of d.files.filter((f) => f.name.toLowerCase().endsWith(".sqlite3"))) {
      items.push(li(x.name, "is-file", () => { $("mAlt").value = x.path; $("mBrowser").hidden = true; }));
    }
    if (d.message) items.push(el("li", "muted maint__bmsg", d.message));
    $("mBList").replaceChildren(...items);
    $("mBrowser").hidden = false;
  } catch (err) { toastError(err); }
}
$("mBrowse").addEventListener("click", () => {
  const v = altText() || $("mCurDb").textContent;
  browse(v.replace(/[\\/][^\\/]*\.sqlite3$/i, ""));
});
$("mBClose").addEventListener("click", () => { $("mBrowser").hidden = true; });
$("mAltClear").addEventListener("click", () => { $("mAlt").value = ""; });
$("mAltForm").addEventListener("submit", (e) => e.preventDefault());

/* ---------- ソリ公差・設備仕様 ---------- */
function sideOut(x, label) {
  if (!x) return [];
  if (x.status !== "ok") return [[label, x.note, "is-warn"]];
  return [[label, x.text]];
}

$("mSori").addEventListener("submit", async (e) => {
  e.preventDefault();
  try {
    const d = await api.post("/api/maint/sori", { atu: $("mSAtu").value, haba: $("mSHaba").value, take: $("mSTake").value, alt: altText() });
    const pairs = sideOut(d.current, "いまの表");
    if (d.current.status === "ok") pairs.push(["使った列", `${d.current.haba_col} / ${d.current.take_col}`]);
    if (d.alt) {
      const diff = d.alt.text !== d.current.text;
      pairs.push(...sideOut(d.alt, "比べる表").map((p) => (diff ? [p[0], p[1], "is-diff"] : p)));
    }
    kv($("mSoriOut"), pairs);
  } catch (err) { toastError(err); }
});

$("mEq").addEventListener("submit", async (e) => {
  e.preventDefault();
  try {
    const d = await api.post("/api/maint/equipment", {
      line: $("mELine").value, zai: $("mEZai").value, atu: $("mEAtu").value,
      haba: $("mEHaba").value, take: $("mETake").value, wei: $("mEWei").value,
    });
    kv($("mEqOut"), d.status === "ok" ? [["結果", d.text || "(空)", d.text === "可" ? "" : "is-warn"]] : [["結果", d.note, "is-warn"]]);
  } catch (err) { toastError(err); }
});

/* ---------- 端末設定のセット ---------- */
function renderTerm(t) {
  kv($("mTerm"), [["設備名", t.facility_text, t.unset ? "is-warn" : ""], ["号機", t.line_no], ["モード", t.mode_text]]);
  $("mTermLv").textContent = t.unset ? "未設定" : t.label;
  $("mTermLv").className = "lv " + (t.unset ? "lv--warn" : "lv--ok");
  $("mPresets").replaceChildren(...t.presets.map((p) => {
    const b = el("button", "btn btn--select", p.label);
    b.type = "button";
    b.title = p.vba;
    b.addEventListener("click", () => setTerm(p.key, `端末設定を「${p.label}」にします。よろしいですか？`));
    return b;
  }));
}

async function setTerm(action, question) {
  if (!(await confirmBox(question, { ok: action === "clear" ? "消す" : "セットする", cancel: "やめる", danger: action === "clear" }))) return;
  try {
    const d = await api.post("/api/maint/terminal", { action });
    renderTerm(d.terminal);
    toast("端末設定を変えました", "ok");
  } catch (err) { toastError(err); }
}
$("mClear").addEventListener("click", () => setTerm("clear", "設備名・号機・モードを消します。次に検査画面を開くと、この端末の設定を求めます。よろしいですか？"));

async function init() {
  try {
    const d = await api.get("/api/maint/state");
    S = d;
    $("mDeny").hidden = d.admin;
    $("maint").dataset.admin = d.admin ? "1" : "0";
    $("mCurDb").textContent = d.quality_db;
    $("mELine").replaceChildren(...d.facilities.map((f) => { const o = el("option", "", f); o.value = f; return o; }));
    if (d.terminal.facility) $("mELine").value = d.terminal.facility;
    renderTerm(d.terminal);
  } catch (err) { toastError(err); }
}

init();
