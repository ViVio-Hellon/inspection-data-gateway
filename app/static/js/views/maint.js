/*
  保守画面(タブ: 歪公差チェック / 仮ロット / マスタ管理 / 測定機器ID / ソリ・設備仕様 / 端末設定・パスワード)。
  誰でも開けて見られる。編集はサーバが 403 {code:"locked"} で断るので、鍵の窓(admin_unlock.js)を出して呼び直す。
  値の書式・状態・違いの有無・入力の誤りはサーバ(ufkensa/kensa/maint.py・master_admin.py・kari_lot.py)が決める。
  ここは描いて送るだけ。
*/

import { api } from "../api.js";
import { confirmBox } from "../askbox.js";
import { adminState, mountLockBadge, withUnlock } from "../admin_unlock.js";
import { toast, toastError } from "../toast.js";
import { testJudgeImage } from "../judgebox.js";
import { tokenUrl } from "../api.js";
import { createMasterGrid } from "./master_grid.js";

const $ = (id) => document.getElementById(id);
const STATUS = { ok: ["", ""], none: ["該当なし", "lv--warn"], error: ["エラー", "lv--ng"], not_found: ["ロットなし", "lv--ng"] };
// 一覧に出す欄(コメントは長いので最後・折り返し)
const LIST_KEYS = ["MG", "HaHiz", "TaHiz", "Naka", "Mimi", "NoTex", "Source", "Flag", "HizCome"];
// 結果カードの大きな札(公差そのもの)と、その下の行
const TILE_KEYS = ["HaHiz", "TaHiz", "Naka", "Mimi", "MG"];
const SUB_KEYS = ["Source", "HizCome", "NoTex", "Flag"];
const TAB_KEY = "maint.tab";

let S = { labels: {}, keys: [] };
let lastAll = null;

// 測定機器ID 画面の状態(タブの初期表示より前に用意する。後ろで宣言すると「初期化前に参照」で落ちる)
const RX = { loaded: false, last: 0, timer: null, lines: [], selected: null, defaults: null, confPort: "" };

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

function kariChip(r) {
  const c = el("span", "lv lv--info maint__kari-chip", "仮ロット");
  c.title = r.source ? `仮ロット(${r.source} から写したもの)` : "仮ロット(手元DB に保存した自分のロット)";
  return c;
}

/* ---------- 履歴出力(テスト履歴・PDF・CSV) ---------- */
// 宣言はタブの初期表示より前に置く(後ろだと「初期化前に参照」で落ちる)
const H = { row: 0 };

async function openHist() {
  try {
    const t = await api.get("/api/maint/history/state");
    $("hTestLot").textContent = t.test_lot;
    $("hTestLv").className = "lv " + (t.test_loaded ? "lv--warn" : "");
    $("hTestLv").textContent = t.test_loaded ? "入っています" : "";
  } catch (err) { toastError(err); }
  await histRows();
}

async function histRows() {
  try {
    const d = await api.get("/api/history/lots");
    $("hCount").textContent = `${d.sheet}: ${d.rows.length} 件`;
    if (!d.rows.some((r) => r.row === H.row)) H.row = 0;
    $("hRows").replaceChildren(...(d.rows.length ? d.rows.map((r) => {
      const tr = el("tr");
      const rb = el("input"); rb.type = "radio"; rb.name = "hPick"; rb.checked = r.row === H.row;
      rb.setAttribute("aria-label", `${r.lot} を選ぶ`);
      const pick = () => { H.row = r.row; rb.checked = true; $("hPdf").disabled = false; };
      rb.addEventListener("change", pick);
      tr.addEventListener("click", pick);
      const td0 = el("td"); td0.append(rb);
      tr.append(td0, el("td", "", r.time), el("td", "", r.lot), el("td", "", r.atu), el("td", "", r.hab), el("td", "", r.tak));
      return tr;
    }) : [(() => { const tr = el("tr"); const td = el("td", "muted", "履歴がありません。テスト履歴を入れると試せます"); td.colSpan = 6; tr.append(td); return tr; })()]));
    $("hPdf").disabled = !H.row;
  } catch (err) { toastError(err); }
}

function histOut(title, files, messages) {
  const items = [el("li", "maint__ftitle", title)];
  for (const f of files || []) items.push(el("li", "", f));
  if (!files || !files.length) items.push(el("li", "muted", "書いたファイルはありません"));
  for (const m of messages || []) if (m.text) items.push(el("li", m.level === "error" || m.level === "warn" ? "maint__fwarn" : "muted", m.text));
  $("hOut").replaceChildren(...items);
}

async function histRun(btn, url, body, title, pick) {
  btn.disabled = true;
  try {
    const r = await api.post(url, body);
    histOut(title, pick(r), r.messages);
    toast(title + " が終わりました", "ok");
  } catch (err) { toastError(err); } finally { btn.disabled = false; }
}

$("hReload").addEventListener("click", histRows);
$("hPdf").addEventListener("click", () => histRun($("hPdf"), "/api/maint/history/pdf", { row: H.row },
  "履歴 PDF 出力", (r) => [...(r.pdf || []), ...(r.xlsx || [])]));
$("hCsv").addEventListener("click", () => histRun($("hCsv"), "/api/maint/history/csv", {}, "CSV 出力", (r) => r.files));
$("hNippou").addEventListener("click", () => histRun($("hNippou"), "/api/maint/history/nippou", {}, "日報 CSV 出力", (r) => r.files));
$("hTestAdd").addEventListener("click", async () => {
  try {
    const r = await api.post("/api/maint/history/test", { add: true });
    toast(r.added.length ? `テスト履歴を入れました(${r.added.join("・")})` : "テスト履歴はもう入っています", r.added.length ? "ok" : "warn");
    await openHist();
  } catch (err) { toastError(err); }
});
$("hTestDel").addEventListener("click", async () => {
  if (!(await confirmBox("テスト履歴(保存用・Lot情報履歴 の行)を消しますか？", { ok: "消す", danger: true }))) return;
  try {
    const r = await api.post("/api/maint/history/test", { add: false });
    toast(r.removed.length ? `消しました(${r.removed.join("・")})` : "テスト履歴は入っていません", r.removed.length ? "ok" : "warn");
    await openHist();
  } catch (err) { toastError(err); }
});

/* ---------- タブ ---------- */
const grids = {};
function showTab(name) {
  if (!document.querySelector(`[data-panel="${name}"]`)) name = "dist";
  for (const t of document.querySelectorAll("#mTabs .tab")) t.setAttribute("aria-selected", String(t.dataset.tab === name));
  for (const p of document.querySelectorAll("[data-panel]")) p.hidden = p.dataset.panel !== name;
  try { localStorage.setItem(TAB_KEY, name); } catch { /* 保存できなくても動く */ }
  if (name === "master") openMaster();
  if (name === "device") openDevice();
  if (name === "hist") openHist();
  if (name === "judge") openJudge();
  if (name === "confi") openConfi();
  if (name === "rstop") openRStop(); else stopRStop();
}
for (const t of document.querySelectorAll("#mTabs .tab")) t.addEventListener("click", () => showTab(t.dataset.tab));

/* ---------- 歪公差: 1 ロットと まとめて の切り替え ---------- */
function setDMode(mode) {
  for (const b of document.querySelectorAll("[data-dmode]")) b.setAttribute("aria-pressed", String(b.dataset.dmode === mode));
  $("dOne").hidden = mode !== "one";
  $("dAll").hidden = mode !== "all";
  $("mResult").replaceChildren(el("p", "maint__empty", mode === "one"
    ? "ロットNo を入れて「引く」を押すと、ここに結果が大きく出ます。"
    : "「まとめて引く」を押すと、ロットごとに 1 行で出します(数が多いと少し待ちます)。"));
  lastAll = null;
}
for (const b of document.querySelectorAll("[data-dmode]")) b.addEventListener("click", () => setDMode(b.dataset.dmode));

function renderAltChip() {
  const a = altText();
  $("mAltChip").hidden = !a;
  $("mAltChip").textContent = a ? `比べる: ${a.split(/[\\/]/).slice(-2).join("/")}` : "";
}

/* ---------- 歪公差: 1 ロット(大きな札。比べるときだけ 変更前 → 変更後) ---------- */
function tileValue(box, a, b, diff) {
  if (b === undefined || !diff) {
    box.append(el("span", "maint__tv", a || "—"));
    return;
  }
  const row = el("span", "maint__tvrow");
  row.append(el("span", "maint__tv maint__tv--old", a || "(空)"), el("span", "maint__arrow", "→"), el("span", "maint__tv maint__tv--new", b || "(空)"));
  box.append(row);
}

function renderOne(d) {
  const r = d.row;
  const cur = r.current, alt = r.alt;
  const card = el("div", "card maint__card");
  const head = el("div", "maint__chead");
  head.append(el("span", "maint__clot", r.lot), ...(r.kari ? [kariChip(r)] : []), statusChip(cur.status));
  if (alt) head.append(el("span", "lv " + (r.diff.length ? "lv--warn" : "lv--ok"), r.diff.length ? `変更後で ${r.diff.length} 項目が変わる` : "変更後も同じ"));
  const meta = el("div", "maint__cmeta");
  meta.append(el("span", "", "検番 "), el("code", "", r.ken_no || "-"), el("span", "", "　引当 "),
              el("span", "", (r.hiki || []).map((h) => h[0]).join(" / ") || "-"));
  card.append(head, meta);

  if (cur.status === "ok" || (alt && alt.status === "ok")) {
    const tiles = el("div", "maint__tiles");
    for (const k of TILE_KEYS) {
      const diff = r.diff.includes(k);
      const t = el("div", "maint__tile" + (diff ? " is-diff" : ""));
      t.append(el("span", "maint__tl", S.labels[k]));
      tileValue(t, cur.values[k], alt ? alt.values[k] : undefined, diff);
      tiles.append(t);
    }
    const sub = el("dl", "maint__csub");
    for (const k of SUB_KEYS) {
      const diff = r.diff.includes(k);
      const dd = el("dd", diff ? "is-diff" : "");
      if (diff) dd.append(el("span", "maint__tv--old", cur.values[k] || "(空)"), el("span", "maint__arrow", " → "), el("b", "", alt.values[k] || "(空)"));
      else dd.textContent = cur.values[k] || "—";
      sub.append(el("dt", "", S.labels[k] + (k === "Flag" ? "(参考)" : "")), dd);
    }
    card.append(tiles, sub);
  }
  for (const [name, x] of [[alt ? "変更前" : "", cur], ["変更後", alt]]) {
    if (!x) continue;
    if (x.trace && x.trace.length) card.append(traceBox(x, name));
    if (x.note) card.append(el("p", "maint__note", pre + x.note));
    for (const m of x.messages || []) card.append(el("p", "maint__msg", pre + m));
  }
  for (const m of r.messages || []) card.append(el("p", "maint__msg", m));
  $("mResult").replaceChildren(card);
}

/** どの表のどの条件で何行残ったか。最初に 0 行になった所(躓いた所)を赤で目立たせる */
function traceBox(x, name) {
  const box = el("div", "maint__trace");
  box.append(el("h4", "maint__trh", `${name ? name + ": " : ""}判定の流れ　参照マスタ: ${x.master || "-"}`));
  const ol = el("ol", "maint__tsteps");
  for (const t of x.trace) {
    const li = el("li", "maint__tstep " + (t.ok === true ? "is-ok" : t.ok === false ? "is-ng" : "is-info"));
    const mark = el("span", "maint__mark", t.ok === true ? "✔" : t.ok === false ? "✖" : "・");
    const body = el("div", "maint__sbody");
    body.append(el("b", "", t.step), el("span", "", "　" + (t.detail || "")));
    if (t.before) {
      const bar = el("div", "maint__bar");
      const fill = el("span", "");
      fill.style.width = Math.max(t.kept ? 2 : 0, Math.round((t.kept / t.before) * 100)) + "%";
      bar.append(fill);
      body.append(bar);
    }
    if (t.ok === false && t.rows && t.rows.length) {
      const ul = el("ul", "maint__miss");
      ul.append(el("li", "muted", "ここで外れた行(表の値):"));
      for (const r of t.rows) ul.append(el("li", "", r));
      body.append(ul);
    }
    li.append(mark, body);
    ol.append(li);
  }
  box.append(ol);
  return box;
}

/* ---------- 歪公差: 全ロット(1 行 1 ロット。違う欄は「変更前 → 変更後」) ---------- */
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
  const sum = el("div", "maint__summary");
  sum.append(
    el("b", "", `${d.count} ロット`), el("span", "muted", `${d.sec} 秒`),
    el("span", "lv lv--ok", `OK ${b.ok}`), el("span", "lv lv--warn", `該当なし ${b.none}`),
    el("span", "lv lv--ng", `エラー ${b.error + b.not_found}`),
    ...(d.alt ? [el("span", "lv " + (d.diffs ? "lv--warn" : "lv--ok"), `変更後で変わるロット ${d.diffs}`)] : []),
    ...((d.kari_used || []).length ? [el("span", "lv lv--info", `仮ロット ${d.kari_used.length}: ${d.kari_used.join(" ")}`)] : []),
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
    const lotTd = el("td", "n", r.lot);
    if (r.kari) { lotTd.append(" ", kariChip(r)); tr.classList.add("is-kari"); }
    tr.append(lotTd, el("td", "n", r.ken_no), st, ...LIST_KEYS.map((k) => cell(r, k)));
    tbody.append(tr);
  }
  t.append(thead, tbody);
  const wrap = el("div", "tablewrap maint__allwrap");
  wrap.append(t);
  $("mResult").replaceChildren(sum, wrap);
}

async function runLot(lot) {
  lot = (lot || "").trim().toUpperCase();
  $("mLot").value = lot;
  try {
    renderOne(await api.post("/api/maint/dist", { lot, alt: altText() }));
    toast("ロード済", "ok");                             // Test:107 MsgBox "ロード済"
  } catch (err) { toastError(err); if (err.field === "lot") $("mLot").select(); }
}

/** 仮ロットの面などから: 歪公差の面へ移って 1 ロット引く。 */
function goLot(lot) {
  showTab("dist");
  setDMode("one");
  runLot(lot);
}

$("mOne").addEventListener("submit", (e) => { e.preventDefault(); runLot($("mLot").value); });
$("mKPick").addEventListener("change", () => { if ($("mKPick").value) { runLot($("mKPick").value); $("mKPick").value = ""; } });

$("mAll").addEventListener("submit", async (e) => {
  e.preventDefault();
  try {
    renderAll(await api.post("/api/maint/dist_all", {
      prefix: $("mPrefix").value.trim(), plate_only: $("mPlate").checked, include_kari: $("mWithKari").checked, alt: altText(),
    }));
  } catch (err) { toastError(err); }
});
$("mOnlyDiff").addEventListener("change", () => { if (lastAll) renderAll(lastAll); });
$("mAlt").addEventListener("input", renderAltChip);

/* ---------- マスタ管理 ---------- */
let M = { sources: [], source: "", table: "", loaded: false };

async function openMaster() {
  if (!grids.master) grids.master = createMasterGrid($("mGrid"));
  if (M.loaded) return;
  M.loaded = true;
  try {
    const d = await api.get("/api/maint/master/sources");
    M.sources = d.sources;
    $("mSources").replaceChildren(...d.sources.map((s) => {
      const b = el("button", "maint__msrc", "");
      b.type = "button";
      b.dataset.key = s.key;
      b.append(el("span", "", s.label), el("span", "maint__mtag" + (s.editable ? "" : " is-ro"), !s.exists ? "無し" : s.editable ? "直せる" : "見るだけ"));
      b.title = s.path || s.error || "場所が決まっていません";
      b.disabled = !s.exists;
      b.addEventListener("click", () => pickSource(s.key));
      return b;
    }));
    const first = d.sources.find((s) => s.exists);
    if (first) pickSource(first.key);
  } catch (err) { M.loaded = false; toastError(err); }
}

async function pickSource(key) {
  M.source = key;
  for (const b of $("mSources").querySelectorAll(".maint__msrc")) b.classList.toggle("is-current", b.dataset.key === key);
  try {
    const d = await api.get(`/api/maint/master/tables?source=${encodeURIComponent(key)}`);
    const items = d.tables.map((t) => {
      const li = el("li");
      const b = el("button", "maint__mtab", "");
      b.type = "button";
      b.dataset.table = t.name;
      b.append(el("span", "", t.name), el("span", "muted", t.rows == null ? "" : String(t.rows)));
      b.addEventListener("click", () => pickTable(t.name));
      li.append(b);
      return li;
    });
    if (d.note) items.push(el("li", "maint__msg", d.note));
    if (!d.tables.length && !d.note) items.push(el("li", "muted maint__small", "表がありません"));
    $("mTables").replaceChildren(...items);
    if (d.tables.length) pickTable(d.tables[0].name);
  } catch (err) { toastError(err); $("mTables").replaceChildren(el("li", "maint__msg", err.message)); }
}

function pickTable(name) {
  M.table = name;
  for (const b of $("mTables").querySelectorAll(".maint__mtab")) b.classList.toggle("is-current", b.dataset.table === name);
  grids.master.open(M.source, name);
}

/* (測定機器ID は下の「受信機・受信記録」にまとめた) */

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
      items.push(li(x.name, "is-file", () => { $("mAlt").value = x.path; $("mBrowser").hidden = true; renderAltChip(); }));
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
$("mAltClear").addEventListener("click", () => { $("mAlt").value = ""; renderAltChip(); });
$("mAltForm").addEventListener("submit", (e) => e.preventDefault());

/* ---------- ソリ公差・設備仕様 ---------- */
function sideOut(x, label) {
  if (!x) return [];
  const src = x.master ? [[label + "の参照マスタ", `${x.master}　表「${x.table}」`]] : [];
  if (x.status !== "ok") return [...src, [label, x.note, "is-warn"]];
  return [...src, [label, x.text]];
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
    kv($("mEqOut"), [["参照マスタ", `${d.master}　表「${d.table}」`],
      d.status === "ok" ? ["結果", d.text || "(空)", d.text === "可" ? "" : "is-warn"] : ["結果", d.note, "is-warn"]]);
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
    const d = await withUnlock(() => api.post("/api/maint/terminal", { action }));
    if (!d) return;                                           // 鍵の窓で「やめる」
    renderTerm(d.terminal);
    toast("端末設定を変えました", "ok");
  } catch (err) { toastError(err); }
}
$("mClear").addEventListener("click", () => setTerm("clear", "設備名・号機・モードを消します。次に検査画面を開くと、この端末の設定を求めます。よろしいですか？"));

/* ---------- 判定の画像: 表示時間・大きさ(PC ごと。2026-11-09) ---------- */
let JD = null;
function jdRead() {
  const out = {};
  for (const k of ["ok", "ng", "diag"]) {
    out[k] = { ms: Math.round(parseFloat($(`jdSec-${k}`).value) * 1000), size: parseInt($(`jdSize-${k}`).value, 10) };
  }
  return out;
}
function jdRender(disp) {
  const box = $("mJudge");
  box.replaceChildren(...["ok", "ng", "diag"].map((k) => {
    const c = el("div", "jd__card");
    c.dataset.kind = k;
    c.append(el("h4", "", JD.labels[k]));
    const r1 = el("label", "jd__row");
    const sec = el("input", "input");
    Object.assign(sec, { type: "number", id: `jdSec-${k}`, min: JD.ms_min / 1000, max: JD.ms_max / 1000, step: 0.5,
                         value: disp[k].ms / 1000 });
    r1.append("表示時間", sec, `秒(${JD.ms_min / 1000}〜${JD.ms_max / 1000}、既定 ${JD.defaults[k].ms / 1000})`);
    const r2 = el("div", "jd__row");
    const size = el("input", "input");
    Object.assign(size, { type: "number", id: `jdSize-${k}`, min: JD.size_min, max: JD.size_max, step: 5, value: disp[k].size });
    size.setAttribute("aria-label", `${JD.labels[k]}の表示サイズ(%)`);
    r2.append("表示サイズ", size, "% ");
    for (const p of JD.presets) {
      const b = el("button", "btn btn--select", p.label);
      b.type = "button";
      b.addEventListener("click", () => { size.value = p.size; });
      r2.append(b);
    }
    const th = el("div", "jd__thumbs");
    const slots = JD.slots.filter((x) => x.kind === k);
    for (const sl of slots) {
      const f = el("figure");
      const im = el("img");
      im.src = tokenUrl(`/api/kensa/judge_image/${encodeURIComponent(sl.slot)}`);
      im.alt = sl.label;
      f.append(im, el("figcaption", sl.source === "missing" ? "is-warn" : "", sl.path ? sl.file : `${sl.label}(既定)`));
      f.title = sl.note || sl.path;
      th.append(f);
    }
    const r3 = el("div", "jd__row");
    let pick = null;
    if (k === "diag") {
      pick = el("select", "input");
      pick.id = "jdSlot-diag";
      pick.append(new Option("乱数(検査と同じ)", ""), ...slots.map((x) => new Option(x.label, x.slot)));
      pick.style.width = "auto";
      r3.append(pick);
    }
    const t = el("button", "btn", "テスト表示");
    t.type = "button";
    t.id = `jdTest-${k}`;
    t.addEventListener("click", () => {
      const v = jdRead()[k];
      testJudgeImage(k, { slot: pick ? pick.value : "", ms: Math.min(Math.max(v.ms || 0, JD.ms_min), JD.ms_max),
                          size: Math.min(Math.max(v.size || 0, JD.size_min), JD.size_max) });
    });
    r3.append(t, el("span", "muted maint__small", "検査画面と同じ窓・大きさ・時間で出します(押すか Esc で消える)"));
    c.append(r1, r2, th, r3);
    return c;
  }));
}
async function openJudge() {
  try {
    JD = await api.get("/api/maint/judge_display");
    jdRender(JD.display);
  } catch (err) { toastError(err); }
}
$("mJudgeSave").addEventListener("click", async () => {
  try {
    const d = await withUnlock(() => api.post("/api/maint/judge_display", { display: jdRead() }));
    if (!d) return;
    JD.display = d.display;
    jdRender(d.display);
    toast("判定の画像の出し方を保存しました(検査画面は 30 秒以内に変わります)", "ok");
  } catch (err) { toastError(err); }
});
$("mJudgeReset").addEventListener("click", () => { if (JD) jdRender(JD.defaults); });

/* ---------- 公差外の確認を出す行 ---------- */
let CR = null;
function crRender(useDefault) {
  $("mConfi").replaceChildren(...CR.kinds.flatMap((k) => {
    const dt = el("dt");
    const lab = el("label", "", k.label);
    lab.htmlFor = `crRows-${k.kind}`;
    dt.append(lab);
    const dd = el("dd");
    const i = el("input", "input");
    Object.assign(i, { id: `crRows-${k.kind}`, value: useDefault ? k.default : k.rows, maxLength: 60 });
    i.dataset.kind = k.kind;
    dd.append(i, el("span", "muted maint__small", ` 1〜${k.max}(既定 ${k.default})`));
    return [dt, dd];
  }));
}
async function openConfi() {
  try {
    CR = await api.get("/api/maint/confi_rows");
    crRender(false);
  } catch (err) { toastError(err); }
}
$("mConfiSave").addEventListener("click", async () => {
  if (!CR) return;
  const rows = Object.fromEntries(CR.kinds.map((k) => [k.kind, $(`crRows-${k.kind}`).value]));
  try {
    const d = await withUnlock(() => api.post("/api/maint/confi_rows", { rows }));
    if (!d) return;
    CR = d;
    crRender(false);
    toast("公差外の確認を出す行を保存しました(全 PC 共通)", "ok");
  } catch (err) { toastError(err); }
});
$("mConfiReset").addEventListener("click", () => { if (CR) crRender(true); });

/* ---------- 仮ロット ---------- */
let K = { fields: [], rows: [], original: "", hikiExtra: [], source: "" };

function kInput(f) {
  const i = el("input", "input" + (f.num ? " input--num" : ""));
  i.id = "mK_" + f.key;
  i.dataset.key = f.key;
  i.maxLength = 100;
  const lab = el("label", "", f.key);
  lab.htmlFor = i.id;
  lab.title = `${f.column}${f.hiki ? " / 引当 " + f.hiki : ""}${f.note ? "\n" + f.note : ""}`;
  return [lab, i];
}

function buildKariForm() {
  $("mKMain").replaceChildren(...K.fields.filter((f) => f.group === "main").flatMap(kInput));
  $("mKDetail").replaceChildren(...K.fields.filter((f) => f.group === "detail").flatMap(kInput));
}

function kValues() {
  return Object.fromEntries(K.fields.map((f) => [f.key, $("mK_" + f.key).value]));
}

function kBody() {
  return { lot: $("mKLot").value.trim().toUpperCase(), ken_head: $("mKHead").value.trim().toUpperCase(), values: kValues(),
           hiki_extra: K.hikiExtra, source: K.source, original: K.original };
}

function kRenderKen() {
  const h = $("mKHead").value.trim().toUpperCase();
  $("mKKen").textContent = /^[A-Z]$/.test(h) ? h + "1".repeat(6) : "-";
}

function kFill({ lot = "", ken_head = "", values = {}, hiki_extra = [], source = "", original = "", notes = [] } = {}) {
  K.original = original;
  K.hikiExtra = hiki_extra || [];
  K.source = source || "";
  $("mKLot").value = lot;
  $("mKHead").value = ken_head;
  for (const f of K.fields) $("mK_" + f.key).value = values[f.key] ?? f.default ?? "";
  $("mKMode").textContent = original ? `仮ロット ${original} を直す` : "新しい仮ロット";
  const extra = K.hikiExtra.length ? [`引当 2 件目から: ${K.hikiExtra.map((h) => h.join(" / ")).join("、")}`] : [];
  $("mKNote").textContent = [...(source ? [`${source} から写しました`] : []), ...notes, ...extra].join("\n");
  kRenderKen();
  for (const x of $("mKForm").querySelectorAll(".is-bad")) x.classList.remove("is-bad");
}

function kMark(err) {
  for (const x of $("mKForm").querySelectorAll(".is-bad")) x.classList.remove("is-bad");
  const target = err.field && $("mKForm").querySelector(`[data-key="${CSS.escape(err.field)}"]`);
  if (target) {
    target.classList.add("is-bad");
    target.closest("details")?.setAttribute("open", "");
    target.focus();
  }
}

function renderKariRows() {
  $("mKCount").textContent = `${K.rows.length} 件(この端末の手元DB)`;
  $("mKariBadge").hidden = !K.rows.length;
  $("mKariBadge").textContent = String(K.rows.length);
  $("mKPick").replaceChildren(el("option", "", K.rows.length ? "(選ぶ)" : "(まだありません)"),
    ...K.rows.map((r) => { const o = el("option", "", `${r.lot}  ${r.values["用途コード"] || ""}`); o.value = r.lot; return o; }));
  $("mKPick").firstChild.value = "";
  if (!K.rows.length) {
    const tr = el("tr");
    const td = el("td", "muted", "まだありません。「新しく作る」か「本物のロットから写す」で作ってください。");
    td.colSpan = 7;
    tr.append(td);
    $("mKRows").replaceChildren(tr);
    return;
  }
  $("mKRows").replaceChildren(...K.rows.map((r) => {
    const v = r.values;
    const tr = el("tr");
    if (r.lot === K.original) tr.className = "is-current";
    const ops = el("td", "maint__kops");
    const btn = (text, cls, fn) => { const b = el("button", "btn btn--sm " + cls, text); b.type = "button"; b.addEventListener("click", fn); return b; };
    ops.append(
      btn("引く", "btn--find", () => goLot(r.lot)),
      btn("直す", "", () => kFill({ ...r, original: r.lot })),
      btn("写す", "", () => kFill({ ...r, lot: "", original: "", notes: [`仮ロット ${r.lot} の値を写しました。ロットNo を入れて保存してください`] })),
      btn("消す", "btn--danger", () => kDelete(r.lot)),
    );
    tr.title = r.updated ? `保存 ${r.updated}${r.source ? " / 元 " + r.source : ""}` : "";
    tr.append(ops, el("td", "n", r.lot), el("td", "n", r.ken_no), el("td", "", `${v["用途コード"] || ""} ${v["用途名"] || ""}`.trim()),
              el("td", "n", `${v["製造材質"] || ""}/${v["製造調質"] || ""}`),
              el("td", "n", `${v["オーダー板厚"]}×${v["オーダー板幅"]}×${v["オーダー板丈"]}`), el("td", "", v["納入先"] || ""));
    return tr;
  }));
}

async function loadKari() {
  const d = await api.get("/api/maint/kari");
  K.fields = d.fields;
  K.rows = d.rows;
  buildKariForm();
  kFill();
  renderKariRows();
}

async function kDelete(lot) {
  if (!(await confirmBox(`仮ロット ${lot} を消します。よろしいですか？`, { ok: "消す", danger: true }))) return;
  try {
    const d = await withUnlock(() => api.post("/api/maint/kari/delete", { lot }));
    if (!d) return;
    K.rows = d.rows;
    if (K.original === lot) kFill();
    renderKariRows();
    toast(`仮ロット ${lot} を消しました`, "ok");
  } catch (err) { toastError(err); }
}

$("mKForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  try {
    const d = await withUnlock(() => api.post("/api/maint/kari", kBody()));
    if (!d) return;
    K.rows = d.rows;
    kFill({ ...d.row, original: d.row.lot });
    renderKariRows();
    toast(`仮ロット ${d.row.lot} を${d.action}しました(検番 ${d.row.ken_no})`, "ok");
  } catch (err) { kMark(err); toastError(err); }
});

$("mKRun").addEventListener("click", async () => {
  try {
    const d = await api.post("/api/maint/kari/dist", { ...kBody(), alt: altText() });
    showTab("dist");
    setDMode("one");
    renderOne(d);
  } catch (err) { kMark(err); toastError(err); }
});

$("mKCopy").addEventListener("click", async () => {
  const src = $("mKSource").value.trim().toUpperCase();
  $("mKSource").value = src;
  try {
    const d = await api.get(`/api/maint/kari/from_real?lot=${encodeURIComponent(src)}`);
    kFill({ ...d, lot: "", original: "" });
    $("mKLot").focus();
  } catch (err) { toastError(err); }
});
$("mKNew").addEventListener("click", () => { kFill(); $("mKLot").focus(); });
$("mKClear").addEventListener("click", () => kFill());
$("mKHead").addEventListener("input", kRenderKen);

/* ---------- 管理者パスワード ---------- */
function renderPw(s) {
  if (!s) return;
  $("mPwLv").textContent = s.custom ? "変更済み" : "既定のまま";
  $("mPwLv").className = "lv " + (s.custom ? "lv--ok" : "lv--warn");
  $("mPwNew").placeholder = `${s.min_length} 文字以上`;
}
window.addEventListener("admin-lock-change", (e) => renderPw(e.detail));

$("mPw").addEventListener("submit", async (e) => {
  e.preventDefault();
  $("mPwErr").textContent = "";
  try {
    const d = await api.post("/api/admin/password", { current: $("mPwCur").value, new: $("mPwNew").value, confirm: $("mPwConf").value });
    for (const id of ["mPwCur", "mPwNew", "mPwConf"]) $(id).value = "";
    renderPw(d.state);
    toast(d.message, "ok");
  } catch (err) {
    $("mPwErr").textContent = err.message;
    const target = { current: "mPwCur", new: "mPwNew", confirm: "mPwConf" }[err.field];
    if (target) $(target).select();
  }
});

/* ---------- 管理者以外停止(remote_stop.py。お願いを出す・解除する・各 PC の様子) ---------- */
const RS = { timer: null };

function renderRStop(d) {
  if (!d) return;
  const req = d.request;
  $("rsLv").textContent = !d.available ? "置き場所に届きません" : req ? "お願いを出しています" : "出していません";
  $("rsLv").className = "lv " + (!d.available ? "lv--ng" : req ? "lv--warn" : "lv--ok");
  kv($("rsInfo"), [
    ["置き場所", d.folder || "-"],
    ["いまのお願い", d.request_text || (d.unreadable || "出ていません"), d.unreadable ? "lv--ng" : ""],
    ["この PC", d.me_pc || "-"],
    ["出せるか", d.password_ok && d.administrator ? "出せます"
      : (!d.administrator ? "この PC のアクセス権限に Administrator がありません" : "管理者パスワード(鍵)を入れると出せます")],
    ...(d.why ? [["届かない理由", d.why, "lv--ng"]] : []),
  ]);
  $("rsIssue").disabled = !d.available || !!req || !d.administrator;
  $("rsClear").disabled = !d.available || !req || !d.administrator;
  const pcs = d.pcs || [];
  $("rsCount").textContent = `${pcs.length} 件`;
  $("rsRows").replaceChildren(...(pcs.length ? pcs.map((p) => {
    const tr = el("tr");
    const cls = p.state === "forced" ? "lv lv--ng" : p.state === "stopped" ? "lv lv--ok" : p.stale ? "lv lv--ng" : "lv lv--warn";
    const st = el("td"); st.append(el("span", req || p.stale ? cls : "lv", p.label));
    tr.append(el("td", "", p.pc), el("td", "", p.login || ""), st, el("td", "", p.detail || ""), el("td", "", p.version || ""), el("td", "", p.beat || ""));
    return tr;
  }) : [(() => { const tr = el("tr"); const td = el("td", "muted", "知らせはまだありません"); td.colSpan = 6; tr.append(td); return tr; })()]));
}

async function loadRStop() {
  try { renderRStop(await api.get("/api/remote_stop/state")); } catch (err) { $("rsMsg").textContent = err.message; }
}
function openRStop() { loadRStop(); if (!RS.timer) RS.timer = setInterval(loadRStop, 5000); }
function stopRStop() { if (RS.timer) { clearInterval(RS.timer); RS.timer = null; } }

async function rsSend(url, body, question) {
  $("rsMsg").textContent = "";
  if (question && !(await confirmBox(question))) return;
  try {
    const d = await withUnlock(() => api.post(url, body));
    renderRStop(d.state);
    toast(d.message, "ok");
  } catch (err) { $("rsMsg").textContent = err.message; }
}
$("rsForm").addEventListener("submit", (e) => {
  e.preventDefault();
  rsSend("/api/remote_stop/issue", { reason: $("rsReason").value.trim() },
    "この PC 以外で動いている UF検査判定 を止めます。ほかの PC では、保存していない入力が残っていても 5 分で強制で止まります。よろしいですか?");
});
$("rsClear").addEventListener("click", () => rsSend("/api/remote_stop/clear", {}, ""));
$("rsReload").addEventListener("click", loadRStop);

async function init() {
  mountLockBadge($("mLock"));
  let tab = "dist";
  try { tab = new URLSearchParams(location.search).get("tab") || localStorage.getItem(TAB_KEY) || "dist"; } catch { /* 既定 */ }
  showTab(tab);
  try {
    const d = await api.get("/api/maint/state");
    S = d;
    $("mCurDb").textContent = d.quality_db;
    $("mELine").replaceChildren(...d.facilities.map((f) => { const o = el("option", "", f); o.value = f; return o; }));
    if (d.terminal.facility) $("mELine").value = d.terminal.facility;
    renderTerm(d.terminal);
    await loadKari();
    renderPw(await adminState());
  } catch (err) { toastError(err); }
}

init();



/* ---------- 測定機器ID: 受信機カード(案B)+ 受信記録 + 段の流れ(案A)+ 当たった行を光らせる(案C) ---------- */
const KIND = { uwave: ["U-WAVE-R", "rx--uw"], ddm: ["DDM-100(コンベックス)", "rx--ddm"], com: ["COM", "rx--com"] };
const RX_CONF = [["baud", "rxBaud"], ["bytesize", "rxBits"], ["parity", "rxParity"], ["stopbits", "rxStop"], ["flow", "rxFlowCtl"]];

function rxCond() {
  return { line: $("mDevLine").value, atu: $("mDevAtu").value, hab: $("mDevHab").value, youto: $("mDevYouto").value,
           st: $("mDevSt").value || "S1", tak: RX.defaults?.tak || "" };
}

async function openDevice() {
  if (!grids.device) grids.device = createMasterGrid($("mDevGrid"), { compact: true });
  if (RX.loaded) { rxPoll(); return; }
  RX.loaded = true;
  grids.device.open("denso", "測定機器ID");
  try {
    RX.defaults = await api.get("/api/maint/device/defaults");
    $("mDevLine").replaceChildren(...RX.defaults.lines.map((x) => { const o = el("option", "", x); o.value = x; return o; }));
    rxFillCond(RX.defaults);
  } catch (err) { toastError(err); }
  rxPoll();
}

function rxFillCond(d) {
  if (d.line) $("mDevLine").value = d.line;
  if (d.atu !== undefined) $("mDevAtu").value = d.atu;
  if (d.hab !== undefined) $("mDevHab").value = d.hab;
  if (d.youto !== undefined) $("mDevYouto").value = d.youto;
}

async function rxPoll() {
  clearTimeout(RX.timer);
  try {
    const q = new URLSearchParams({ since: String(RX.last), ...rxCond() });
    const d = await api.get(`/api/maint/rx?${q}`);
    rxCards(d.cards, d.ports);
    if (d.lines.length) {
      RX.lines.unshift(...d.lines.reverse());
      RX.lines.length = Math.min(RX.lines.length, 200);
      RX.last = d.last;
      rxRows();
      rxSelect(RX.lines[0]);                      // 届いた最新の値を流れに出す
    }
    if (d.cards.some((c) => c.running && c.where === "mnt")) RX.timer = setTimeout(rxPoll, 500);
  } catch (e) { RX.timer = setTimeout(rxPoll, 2000); }
}

function rxCards(cards, ports) {
  const box = $("rxCards");
  RX.kinds = Object.fromEntries(cards.map((c) => [c.port, c.kind]));
  const items = cards.map((c) => {
    const [kname, kcls] = KIND[c.kind] || KIND.com;
    const card = el("div", `card rx__card ${kcls}`);
    const head = el("div", "rx__chead");
    const st = c.running && c.where === "mnt" ? ["受信中", "lv--ok"] : c.running && c.where === "kensa" ? ["検査画面で受信中", "lv--info"]
      : c.error ? ["開けません", "lv--ng"] : ["止まっています", ""];
    head.append(el("b", "rx__kname", kname), el("span", "lv " + st[1], st[0]));
    const sub = el("div", "rx__port", c.kind === "uwave" && c.port.startsWith("FTDI:") ? `USB 直接(シリアル ${c.port.slice(5)})` : c.label);
    const big = el("div", "rx__big", c.last ? (c.last.test?.value || c.last.text) : "—");
    const meta = el("div", "rx__meta", c.last ? `最後 ${c.last.at}・この画面で ${c.count} 件` : (c.present ? "まだ受け取っていません" : "このポートが見つかりません(抜けている?)"));
    const items = [head, sub, big, meta];
    if (c.error) {
      const busy = /使って|使用中/.test(c.error);
      items.push(el("p", "rx__err", busy ? `${c.error}\n→ ${c.kind === "ddm" ? "Bluetooth キー入力ユーティリティ" : "U-WAVEPAK"} を終了してから、もう一度「受信を始める」。` : c.error));
    }
    const ops = el("div", "rx__ops");
    const go = el("button", "btn btn--sm " + (c.running && c.where === "mnt" ? "" : "btn--find"), c.running && c.where === "mnt" ? "止める" : "受信を始める");
    go.type = "button";
    go.addEventListener("click", () => rxStartStop(c));
    const kz = el("label", "rx__kensa");
    const chk = el("input"); chk.type = "checkbox"; chk.checked = c.kensa;
    chk.addEventListener("change", () => rxKensa(c, chk.checked, chk));
    kz.append(chk, document.createTextNode(" 検査画面はこれだけ使う"));
    // 2026-11-09: 印が 1 つも無ければ、ここに足した受信機を全部 検査・校正画面で使う。印を付けると印の付いたものだけ
    kz.title = "印が無いときは、ここに足した受信機を全部 検査・校正画面で使います。印を付けると、印の付いた受信機だけを使います";
    ops.append(go, kz);
    if (c.kind !== "ddm") {
      const cf = el("button", "btn btn--sm", "設定"); cf.type = "button";
      cf.title = `${c.conf.baud}bps・${c.conf.bytesize}${c.conf.parity}${c.conf.stopbits}・${c.conf.flow}`;
      cf.addEventListener("click", () => rxOpenConf(c));
      ops.append(cf);
    }
    const rm = el("button", "btn btn--sm rx__rm", "外す"); rm.type = "button";
    rm.title = "この受信機をこの画面から外す(検査画面でも受け取らなくなる)";
    rm.addEventListener("click", () => rxRemove(c));
    ops.append(rm);
    items.push(ops);
    card.append(...items);
    return card;
  });
  // 足す: 一覧から(U-WAVE-R の直接・Bluetooth の COM・USB の COM)
  // 2026-11-09: 0.5 秒ごとの描き直しで選択が勝手に戻っていた。選んでいる最中は作り直さず、選んだ値は引き継ぐ
  const oldAdd = box.querySelector(".rx__add");
  const act = document.activeElement;
  if (oldAdd && act && oldAdd.contains(act)) { box.replaceChildren(...items, oldAdd); return; }
  const prevSel = $("rxAddSel")?.value || "";
  const prevTxt = $("rxAddText")?.value || "";
  const add = el("div", "card rx__add");
  const have = new Set(cards.map((c) => c.port));
  const sel = el("select", "input"); sel.id = "rxAddSel";
  const free = (ports || []).filter((p) => !have.has(p.port));
  sel.append(...(free.length ? free : [{ port: "", label: "(足せる受信機がありません)" }]).map((p) => new Option(p.label || p.port, p.port)));
  if (prevSel && [...sel.options].some((o) => o.value === prevSel)) sel.value = prevSel;
  const txt = el("input", "input"); txt.placeholder = "一覧に無ければ COM5 など"; txt.id = "rxAddText"; txt.value = prevTxt;
  const b = el("button", "btn btn--sm btn--find", "＋ 受信機を足す"); b.type = "button";
  b.addEventListener("click", async () => {
    const port = txt.value.trim() || sel.value;
    try { const d = await api.post("/api/maint/rx/add", { port }); b.blur(); txt.value = ""; rxCards(d.cards, d.ports); } catch (e) { toastError(e); }
  });
  add.append(el("b", "", "受信機を足す"), sel, txt, b);
  box.replaceChildren(...items, add);
}

async function rxStartStop(c) {
  try {
    const on = c.running && c.where === "mnt";
    const r = await api.post(on ? "/api/maint/rx/stop" : "/api/maint/rx/start", { port: c.port });
    if (!on) toast(`${c.port} の受信を始めました。測定器のボタンを押してください`, "ok");
    if (r && r.warning) toast(r.warning, "warn");
  } catch (e) { toastError(e); }
  rxPoll();
}

async function rxKensa(c, on, chk) {
  try {
    await api.post("/api/maint/com/kensa", { on, port: c.port, ...c.conf });
    toast(on ? `検査画面は印の付いた受信機(${c.port} など)だけを使います` : `${c.port} の印を外しました(印が 1 つも無ければ、足した受信機を全部使います)`, "success");
  } catch (e) { chk.checked = !on; toastError(e); }
  rxPoll();
}

async function rxRemove(c) {
  if (!(await confirmBox(`${c.port} をこの画面から外しますか?(検査画面でも受け取らなくなります)`, { ok: "外す", cancel: "やめる" }))) return;
  try { await api.post("/api/maint/rx/remove", { port: c.port }); } catch (e) { toastError(e); }
  rxPoll();
}

function rxOpenConf(c) {
  RX.confPort = c.port;
  $("rxConfTitle").textContent = `受信の設定(${c.port})`;
  for (const [k, id] of RX_CONF) $(id).value = String(c.conf[k]);
  $("rxConfDlg").showModal();
}
$("rxConfDefault").addEventListener("click", () => {
  const d = { baud: 57600, bytesize: 8, parity: "N", stopbits: 1, flow: "rtscts" };
  for (const [k, id] of RX_CONF) $(id).value = String(d[k]);
});
$("rxConfSave").addEventListener("click", async () => {
  const body = { port: RX.confPort };
  for (const [k, id] of RX_CONF) body[k] = $(id).value;
  try { await api.post("/api/maint/rx/conf", body); $("rxConfDlg").close(); toast("受信の設定を保存しました(次に受信を始めたときから)", "success"); } catch (e) { toastError(e); }
  rxPoll();
});

/** 判定の札(受信記録の行・流れの見出し) */
function rxVerdict(t) {
  if (!t || t.error) return [t?.error || "テストできません", "lv--ng"];
  if (!t.accepted) return ["受け付けない", "lv--ng"];
  if (t.kind !== "DT") return ["値は入る", "lv--info"];
  if (!(t.hits || []).length) return ["NG 測定機器ID に無い", "lv--ng"];
  if (t.judge.ok === true) return ["OK " + ((t.hits[0] || {})["機器名"] || ""), "lv--ok"];
  if (t.judge.ok === false) return ["NG 止まる", "lv--ng"];
  return [t.expect ? "機器を確かめられない" : "厚を入れると判定", "lv--warn"];
}

function rxRows() {
  const body = $("rxRows");
  body.replaceChildren(...RX.lines.map((ln) => {
    const tr = el("tr", RX.selected === ln ? "is-sel" : "");
    tr.tabIndex = 0;
    const v = rxVerdict(ln.test);
    const who = el("td"); who.append(el("span", "rx__dot " + kindClass(ln.port)), document.createTextNode(ln.port.startsWith("FTDI:") ? "U-WAVE-R" : ln.port));
    const ver = el("td"); ver.append(el("span", "lv " + v[1], v[0]));
    tr.append(el("td", "", ln.at), who, el("td", "rx__mono", ln.text), el("td", "rx__mono", ln.test?.value || ""), ver);
    tr.addEventListener("click", () => rxSelect(ln));
    tr.addEventListener("keydown", (e) => { if (e.key === "Enter") rxSelect(ln); });
    return tr;
  }));
}
function kindClass(port) { return (KIND[(RX.kinds || {})[port]] || KIND.com)[1]; }

/** 案A: 受信機 → 文字列 → 入る値 → 測定機器ID の行 → 検査の判定。止まった段を赤く */
function rxSelect(ln) {
  RX.selected = ln;
  rxRows();
  const t = ln.test || {};
  const v = rxVerdict(t);
  $("rxDetailTitle").textContent = `判定の流れ(${ln.at})`;
  $("rxVerdict").textContent = v[0];
  $("rxVerdict").className = "lv " + v[1];
  const dt = t.kind === "DT";
  const hit = (t.hits || [])[0];
  const stages = [
    ["受信機", ln.port.startsWith("FTDI:") ? "U-WAVE-R(USB 直接)" : ln.port, true, ""],
    ["届いた文字列", ln.text, true, ln.hex],
    ["入る値", t.accepted ? t.value : "受け付けない", !!t.accepted, t.accepted ? `種別 ${t.kind}${dt ? `・機器番号 ${t.code}` : ""}` : "頭が DT / #L / #〆 ではない"],
    ["測定機器ID の行", !dt ? "(機器の番号なし)" : hit ? `${hit["ライン"]} ${hit["機器名"]}` : "行がありません", !dt ? null : !!hit,
     !dt ? "#L は前の機器のまま" : hit ? `ID ${hit["ID"]}・厚 ${hit["厚下"] || "-"}〜${hit["厚上"] || "-"}` : `ライン ${t.line} に 出力テキスト ${t.code} が無い`],
    ["検査の判定", !t.accepted ? "—" : !dt ? "値は入る" : t.judge.ok === true ? "OK 値が入る" : t.judge.ok === false ? "NG 止まる" : "確かめられない",
     !t.accepted ? false : !dt ? null : t.judge.ok,
     t.judge?.text || (t.expect ? `この厚・用途で使う機器 ${t.expect.name || "(該当なし)"}` : "上の「テストの条件」に厚を入れると、正しい機器かまで確かめます")],
  ];
  let stopped = false;
  $("rxFlow").replaceChildren(...stages.map(([name, val, ok, sub], i) => {
    const st = stopped ? "is-skip" : ok === false ? "is-ng" : ok === true ? "is-ok" : "is-info";
    if (ok === false) stopped = true;
    const li = el("li", "rx__stage " + st);
    li.append(el("small", "", `${i + 1} ${name}`), el("b", "", val || "—"), el("span", "", sub || ""));
    return li;
  }));
  const kv = $("rxKv");
  kv.replaceChildren();
  const pair = (k, x) => kv.append(el("dt", "", k), el("dd", "", x || "—"));
  pair("入力先", t.st ? `${t.st} … ${t.st_label}` : "");
  pair("ほかのラインで同じ番号", (t.others || []).map((r) => `${r["ライン"]} ${r["機器名"]}`).join("・"));
  pair("帳票の機器欄", (t.report || []).join(" / "));
  $("rxNotes").replaceChildren(...(t.notes || []).map((n) => el("p", "maint__note", n)));
  // 案C: 下の表で当たった行を光らせる(このラインの同じ番号 = 判定の色、使うはずの機器 = 枠)
  if (grids.device) {
    grids.device.highlight((r) => {
      if (!dt) return null;
      const sameLine = String(r["ライン"] || "") !== "" && (t.hits || []).some((h) => h["ID"] === r["ID"] && h["機器名"] === r["機器名"] && h["ライン"] === r["ライン"]);
      if (sameLine) return t.judge.ok === false ? "ng" : "ok";
      if (t.expect && t.expect.text && String(r["出力テキスト"]) === String(t.expect.text) && (t.hits[0] ? r["ライン"] === t.hits[0]["ライン"] : true)) return "expect";
      return null;
    });
  }
}

$("rxCond").addEventListener("submit", async (e) => {
  e.preventDefault();
  const text = $("mDevText").value.trim();
  if (!text) { $("mDevText").focus(); return; }
  try {
    const t = await api.post("/api/maint/device/test", { text, ...rxCond() });
    const ln = { g: 0, port: "貼り付け", at: new Date().toLocaleTimeString("ja-JP"), text, hex: "", test: t };
    RX.lines.unshift(ln);
    rxRows();
    rxSelect(ln);
  } catch (err) { toastError(err); if (err.field === "st") $("mDevSt").select(); }
});
$("mDevFromKensa").addEventListener("click", async () => {
  try {
    RX.defaults = await api.get("/api/maint/device/defaults");
    rxFillCond(RX.defaults);
    toast(RX.defaults.lot ? `検査中のロット ${RX.defaults.lot} の値を入れました` : "検査中のロットがありません(この端末のラインだけ入れました)", RX.defaults.lot ? "ok" : "warn");
  } catch (err) { toastError(err); }
});
