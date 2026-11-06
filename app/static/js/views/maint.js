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

// COM受信テストの状態(タブの初期表示より前に用意する。後ろで宣言すると「初期化前に参照」で落ちる)
let comLast = 0, comTimer = null, comLoaded = false, comKensa = [];
const COM_FIELDS = [["baud", "mComBaud"], ["bytesize", "mComBits"], ["parity", "mComParity"], ["stopbits", "mComStop"], ["flow", "mComFlow"]];

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

/* ---------- タブ ---------- */
const grids = {};
function showTab(name) {
  if (!document.querySelector(`[data-panel="${name}"]`)) name = "dist";
  for (const t of document.querySelectorAll("#mTabs .tab")) t.setAttribute("aria-selected", String(t.dataset.tab === name));
  for (const p of document.querySelectorAll("[data-panel]")) p.hidden = p.dataset.panel !== name;
  try { localStorage.setItem(TAB_KEY, name); } catch { /* 保存できなくても動く */ }
  if (name === "master") openMaster();
  if (name === "device") { openDevice(); comOpen(); }
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

/* ---------- 測定機器ID(伝送テスト + 表) ---------- */
let devLoaded = false;
let devDefaults = null;

async function openDevice() {
  if (!grids.device) grids.device = createMasterGrid($("mDevGrid"), { compact: true });
  if (devLoaded) return;
  devLoaded = true;
  grids.device.open("denso", "測定機器ID");
  try {
    devDefaults = await api.get("/api/maint/device/defaults");
    $("mDevLine").replaceChildren(...devDefaults.lines.map((x) => { const o = el("option", "", x); o.value = x; return o; }));
    fillDevice(devDefaults);
  } catch (err) { toastError(err); }
  $("mDevText").focus();
}

function fillDevice(d) {
  if (d.line) $("mDevLine").value = d.line;
  if (d.atu !== undefined) $("mDevAtu").value = d.atu;
  if (d.hab !== undefined) $("mDevHab").value = d.hab;
  if (d.youto !== undefined) $("mDevYouto").value = d.youto;
}

function devRow(label, value, cls) {
  const tr = el("tr", cls || "");
  const td = el("td");
  if (value instanceof Node) td.append(value); else td.textContent = value ?? "";
  tr.append(el("th", "", label), td);
  return tr;
}

function renderDevice(d) {
  const box = $("mDevOut");
  const head = el("div", "maint__devhead");
  const ok = d.judge.ok;
  const verdict = !d.accepted ? ["受け付けない", "lv--ng"] : d.kind !== "DT" ? ["値は入る(機器の確かめなし)", "lv--info"]
    : ok === true ? ["OK 値が入る", "lv--ok"] : ok === false ? ["NG 止まる", "lv--ng"] : ["機器を確かめられない", "lv--warn"];
  head.append(el("span", "maint__devval", d.accepted ? d.value : "—"), el("span", "lv " + verdict[1], verdict[0]));
  const t = el("table", "table table--dense maint__devtbl");
  const tb = el("tbody");
  const hit = (rows) => rows.map((r) => `${r["ライン"]} ${r["機器名"]}(ID ${r["ID"]}・厚 ${r["厚下"] || "-"}〜${r["厚上"] || "-"}${r["用途"] ? "・用途あり" : ""})`).join("\n");
  tb.append(
    devRow("種別", d.kind || "(なし)"),
    devRow("入る値", d.accepted ? `${d.value}(伝送の値部分「${d.raw}」)` : "—"),
    devRow("入力先", `${d.st} … ${d.st_label}`),
    devRow("機器の番号", d.code || "(DT のときだけ。3〜7 文字目)"),
    devRow(`この機器(${d.line})`, d.kind === "DT" ? (hit(d.hits) || "測定機器ID に行がありません") : "—", d.kind === "DT" && !d.hits.length ? "is-warn" : ""),
    ...(d.others.length ? [devRow("ほかのラインで同じ番号", hit(d.others))] : []),
    devRow("この厚・用途で使う機器", d.expect ? `${d.expect.name || "(該当なし)"}(出力テキスト ${d.expect.text || "-"})` : "—"),
    devRow("検査画面の判定", d.judge.text || (ok === true ? "そのまま入って判定に進みます" : "—"), ok === false ? "is-warn" : ""),
    devRow("帳票の機器欄", (d.report || []).join(" / ") || "—"),
  );
  t.append(tb);
  box.replaceChildren(head, t, ...d.notes.map((n) => el("p", "maint__note", n)));
}

async function runDevice() {
  try {
    renderDevice(await api.post("/api/maint/device/test", {
      text: $("mDevText").value, line: $("mDevLine").value, atu: $("mDevAtu").value, hab: $("mDevHab").value,
      youto: $("mDevYouto").value, st: $("mDevSt").value, tak: devDefaults?.tak || "",
    }));
  } catch (err) { toastError(err); if (err.field === "st") $("mDevSt").select(); }
}
$("mDevForm").addEventListener("submit", (e) => { e.preventDefault(); runDevice(); });
$("mDevFromKensa").addEventListener("click", async () => {
  try {
    devDefaults = await api.get("/api/maint/device/defaults");
    fillDevice(devDefaults);
    toast(devDefaults.lot ? `検査中のロット ${devDefaults.lot} の値を入れました` : "検査中のロットがありません(この端末のラインだけ入れました)", devDefaults.lot ? "ok" : "warn");
  } catch (err) { toastError(err); }
});

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


/* ---------- COM受信テスト ---------- */
function comFill(conf) {
  for (const [k, id] of COM_FIELDS) if (conf[k] !== undefined) $(id).value = String(conf[k]);
}
function comPorts(ports, want) {
  const sel = $("mComPort");
  sel.replaceChildren(...(ports.length ? ports : [{ port: "", device: "" }]).map((p) =>
    new Option(p.port ? (p.label || `${p.port}${p.usb ? "(USB の仮想COM)" : ""}`) : "(見つかりません。右に入れてください)", p.port)));
  if (want && ports.some((p) => p.port === want)) sel.value = want;
  else if (want) $("mComPortText").value = want;
}
function comShow(d) {
  $("mComLv").textContent = d.running ? `受信中 ${d.port}(${d.bytes} バイト)` : "止まっています";
  $("mComLv").className = "lv " + (d.running ? "lv--ok" : "");
  $("mComErr").hidden = !d.error;
  $("mComErr").textContent = d.error || "";
  const fresh = (d.lines || []).filter((ln) => ln.seq > comLast);   // 表示済みの行は足さない(止めたときに全部返ってくる)
  if (fresh.length) {
    const body = $("mComRows");
    if (comLast === 0) body.replaceChildren();
    for (const ln of fresh) {
      const use = el("button", "btn btn--sm", "この文字列でテスト");
      use.type = "button";
      use.addEventListener("click", () => { $("mDevText").value = ln.text; runDevice(); });
      const tr = el("tr", "");
      const tdUse = el("td", ""); tdUse.append(use);
      tr.append(el("td", "", ln.at), el("td", "maint__comtext", ln.text), el("td", "maint__comhex", ln.hex), tdUse);
      body.prepend(tr);
    }
    comLast = d.last;
  }
  if (d.kensa) comKensa = d.kensa;
  comKensaShow();
  clearTimeout(comTimer);
  if (d.running) comTimer = setTimeout(comPoll, 500);
}
async function comPoll() {
  try { comShow(await api.get(`/api/maint/com?since=${comLast}`)); } catch (e) { /* 次の回で */ }
}
async function comOpen() {
  if (comLoaded) return;
  comLoaded = true;
  try {
    const d = await api.get("/api/maint/com");
    comPorts(d.ports, d.saved.port);
    comFill(d.saved);
    comLast = d.last || 0;                      // 開く前に受け取った分は出さない
    comShow(d);
  } catch (e) { toastError(e); }
}
$("mComRefresh").addEventListener("click", async () => {
  const d = await api.get("/api/maint/com");
  comPorts(d.ports, $("mComPort").value || $("mComPortText").value);
});
$("mComDefault").addEventListener("click", async () => { const d = await api.get("/api/maint/com"); comFill(d.defaults); });
$("mComForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const body = { port: $("mComPortText").value.trim() || $("mComPort").value };
  for (const [k, id] of COM_FIELDS) body[k] = $(id).value;
  try {
    comLast = 0;
    $("mComRows").replaceChildren(el("tr", "", ""));
    comShow(await api.post("/api/maint/com/start", body));
  } catch (err) { toastError(err); }
});
$("mComStop2").addEventListener("click", async () => {
  try {
    comShow(await api.post("/api/maint/com/stop", {}));
    const d = await api.get("/api/maint/com");
    comPorts(d.ports, d.saved.port);           // 止めたあとも選んでいたポートを残す
  } catch (e) { toastError(e); }
});

$("mComKensa").addEventListener("change", async () => {
  const body = { on: $("mComKensa").checked, port: $("mComPortText").value.trim() || $("mComPort").value };
  for (const [k, id] of COM_FIELDS) body[k] = $(id).value;
  try {
    const r = await api.post("/api/maint/com/kensa", body);
    comKensa = r.kensa || [];
    comKensaShow();
    toast(body.on ? `検査画面を開くと ${body.port} から受け取ります` : `${body.port} は検査画面で受け取りません`, "success");
  } catch (e) { $("mComKensa").checked = !body.on; toastError(e); }
});

function comKensaShow() {
  const cur = $("mComPortText").value.trim() || $("mComPort").value;
  $("mComKensa").checked = comKensa.some((x) => x.port === cur);
  $("mComKensaList").textContent = comKensa.length
    ? "検査画面で受け取る受信機: " + comKensa.map((x) => x.port).join("・")
    : "検査画面で受け取る受信機: なし(キー入力で受け取ります)";
}
$("mComPort").addEventListener("change", comKensaShow);
$("mComPortText").addEventListener("input", comKensaShow);
