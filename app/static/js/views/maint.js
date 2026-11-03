/*
  保守画面。歪公差チェック(仮ロットも)・ソリ公差・設備仕様・端末設定のセット・仮ロット・管理者パスワード。
  誰でも開けて見られる。編集はサーバが 403 {code:"locked"} で断るので、鍵の窓(admin_unlock.js)を出して呼び直す。
  値の書式・状態・違いの有無・入力の誤りはサーバ(ufkensa/kensa/maint.py・kari_lot.py)が決める。ここは描いて送るだけ。
*/

import { api } from "../api.js";
import { confirmBox } from "../askbox.js";
import { adminState, mountLockBadge, withUnlock } from "../admin_unlock.js";
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

function kariChip(r) {
  const c = el("span", "lv lv--info maint__kari-chip", "仮ロット");
  c.title = r.source ? `仮ロット(${r.source} から写したもの)` : "仮ロット(手元DB に保存した自分のロット)";
  return c;
}

/* ---------- 歪公差: 1 ロット(縦に項目、横に いまの表 / 比べる表) ---------- */
function renderOne(d) {
  const r = d.row;
  const cur = r.current, alt = r.alt;
  $("mSummary").replaceChildren(
    el("b", "", r.lot), ...(r.kari ? [kariChip(r)] : []), el("span", "muted", "検番"), el("code", "", r.ken_no || "-"),
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
    renderAll(await api.post("/api/maint/dist_all", {
      prefix: $("mPrefix").value.trim(), plate_only: $("mPlate").checked, include_kari: $("mWithKari").checked, alt: altText(),
    }));
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
      btn("引く", "btn--find", () => { $("mLot").value = r.lot; $("mOne").requestSubmit(); }),
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
    renderOne(await api.post("/api/maint/kari/dist", { ...kBody(), alt: altText() }));
    $("mResult").scrollIntoView({ block: "nearest" });
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
