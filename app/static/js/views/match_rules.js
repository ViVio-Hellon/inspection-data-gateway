/* 保守 → 照合ルール(app/routes/match_rules.py)。
  ルール群ごとにルールを表で直す。表の列・ロットの項目は在るものから選ぶだけ。
  試し照合: ロットNo で、ルールの順に残った行(管理番号・行)と結果を、保存済み と 編集中 で並べる。
  保存: 先に仕掛のロットで結果が変わるものを出す。変わるなら理由が要る(サーバが断る)。鍵は withUnlock。
*/
import { api } from "../api.js";
import { withUnlock } from "../admin_unlock.js";
import { toast, toastError } from "../toast.js";

const $ = (id) => document.getElementById(id);
const R = { state: null, name: "", work: null };

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text ?? "";
  return e;
}
function select(options, value, label) {
  const s = el("select", "input");
  s.setAttribute("aria-label", label);
  const opts = options.includes(value) || value === "" ? options : [value, ...options];
  for (const o of opts) {
    const op = el("option", "", o === value && !options.includes(value) ? `${o}(表に無い)` : o);
    op.value = o;
    s.append(op);
  }
  s.value = value;
  return s;
}
const G = () => R.state.groups.find((g) => g.name === R.name);

function renderGroups() {
  $("rGroups").replaceChildren(...R.state.groups.map((g) => {
    const b = el("button", "btn", g.name);
    b.type = "button";
    b.setAttribute("aria-pressed", String(g.name === R.name));
    b.addEventListener("click", () => { pick(g.name); });
    return b;
  }));
}

function pick(name) {
  R.name = name;
  R.work = structuredClone(G());
  renderGroups();
  render();
  $("rTrial").replaceChildren();
  $("rReg").replaceChildren();
  renderInputs();
}

/** 試し照合の直接入力(測定機器ID: ライン・厚・用途 / 設備仕様: ライン・オーダー厚)。空ならロットの値 */
function renderInputs() {
  const names = G().inputs || [];
  const box = $("rInputs");
  box.hidden = !names.length;
  box.replaceChildren(...names.flatMap((n) => {
    const id = "rIn_" + n;
    const lab = el("label", "", n);
    lab.htmlFor = id;
    const inp = el("input", "input");
    Object.assign(inp, { id, size: n === "ライン" ? 6 : 5, maxLength: 40, placeholder: n === "ライン" ? "全部" : "ロットの値" });
    inp.dataset.input = n;
    return [lab, inp];
  }));
}
function trialInputs() {
  const out = {};
  for (const i of $("rInputs").querySelectorAll("input[data-input]")) if (i.value.trim()) out[i.dataset.input] = i.value.trim();
  return out;
}

function colCount(cmp, now) { return cmp === "範囲" ? (now === 2 ? 2 : 3) : 1; }

function render() {
  const g = G(), w = R.work;
  $("rInfo").textContent = `表「${g.table}」(${g.source}・${g.table_rows} 行)。照合ルールの版 ${g.version || 0}${g.version ? "" : "(既定)"}。${g.note || ""}`;
  $("rHelp").textContent = R.state.help_switch || "";
  $("rSwapRow").hidden = !(g.dim_swap && g.dim_swap.length);
  const src = g.field_sources || {};
  const fieldLabel = (f) => (src[f] ? `${f}【${src[f]}】` : f);
  const use = $("rUse");
  use.replaceChildren(...R.state.use_rows.map((u) => { const o = el("option", "", u); o.value = u; return o; }));
  use.value = w.use_rows;
  $("rNew").checked = !!w.new_way;
  const cols = ["", ...g.table_fields];
  $("rRules").replaceChildren(...w.rules.map((r, i) => {
    const tr = el("tr");
    const ord = el("input", "input input--num");
    Object.assign(ord, { value: r.order, size: 4, inputMode: "numeric" });
    ord.setAttribute("aria-label", "順番");
    ord.addEventListener("change", () => { r.order = Number(ord.value); });
    const colTd = el("td");
    const n = colCount(r.compare, r.columns.length);
    while (r.columns.length < n) r.columns.push("");
    r.columns.length = n;
    r.columns.forEach((c, k) => {
      const s = select(cols, c, ["列", "下の列", "上の列", "値の列"][n > 1 ? k + 1 : 0]);
      s.addEventListener("change", () => { r.columns[k] = s.value; });
      colTd.append(s);
    });
    const fld = select(g.fields, r.field, "ロットの項目");
    for (const op of fld.options) if (g.fields.includes(op.value)) op.textContent = fieldLabel(op.value);
    fld.addEventListener("change", () => { r.field = fld.value; render(); });
    const fldWrap = el("span");
    const badge = el("span", "lv", src[r.field] || "-");
    badge.title = "ロットの項目の出どころ";
    badge.dataset.src = src[r.field] || "";
    if (src[r.field] === "オーダー") badge.className = "lv lv--warn";
    fldWrap.append(badge, " ", fld);
    const cmp = select(R.state.compares, r.compare, "比べ方");
    cmp.addEventListener("change", () => { r.compare = cmp.value; render(); });
    const prep = select(R.state.preps, r.prep, "前処理");
    prep.addEventListener("change", () => { r.prep = prep.value; });
    const blank = select(R.state.blanks, r.blank, "空のとき");
    blank.addEventListener("change", () => { r.blank = blank.value; });
    const en = el("input");
    en.type = "checkbox"; en.checked = !!r.enabled; en.setAttribute("aria-label", "有効");
    en.addEventListener("change", () => { r.enabled = en.checked; });
    const note = el("input", "input");
    Object.assign(note, { value: r.note || "", maxLength: 200, size: 10 });
    note.setAttribute("aria-label", "備考");
    note.addEventListener("change", () => { r.note = note.value; });
    const del = el("button", "btn btn--danger", "消す");
    del.type = "button";
    del.addEventListener("click", () => { w.rules.splice(i, 1); render(); });
    for (const x of [ord, colTd, fldWrap, cmp, prep, blank, en, note, del]) {
      if (x.tagName === "TD") tr.append(x); else { const td = el("td"); td.append(x); tr.append(td); }
    }
    return tr;
  }));
  $("rHist").replaceChildren(...(R.state.history.filter((h) => h.group === R.name).map((h) =>
    el("li", "", `${h.t}${h.version ? " 版" + h.version : ""} ${h.user || ""}@${h.pc || ""}${h.reason ? "(" + h.reason + ")" : ""}: ${(h.summary || []).join("、") || "変更なし"}`)))
    .concat([]).slice(0, 30));
  if (!$("rHist").children.length) $("rHist").append(el("li", "", "(まだ変更はありません)"));
}

$("rUse").addEventListener("change", () => { R.work.use_rows = $("rUse").value; });
$("rNew").addEventListener("change", () => { R.work.new_way = $("rNew").checked; });
$("rAdd").addEventListener("click", () => {
  const max = Math.max(0, ...R.work.rules.map((r) => r.order));
  R.work.rules.push({ order: max + 10, columns: [""], field: G().fields[0], compare: "条件式", prep: "なし", blank: "通す", enabled: true, note: "" });
  render();
});
$("rSwap").addEventListener("click", async () => {
  const pairs = G().dim_swap || [];
  const m = new Map();
  for (const [a, b] of pairs) { m.set(a, b); m.set(b, a); }
  const changed = [];
  for (const r of R.work.rules) {
    if (m.has(r.field)) { changed.push(`順番${r.order}: ${r.field} → ${m.get(r.field)}`); r.field = m.get(r.field); }
  }
  render();
  $("rSwapNote").textContent = changed.length ? `切り替えました(まだ保存していません): ${changed.join("、")}` : "寸法のルールがありません";
  try { await regression(); } catch (err) { toastError(err); }
});
$("rReset").addEventListener("click", () => { R.work = { ...structuredClone(G().default), new_way: true }; render(); });

function rowsText(rows, count) {
  const s = rows.map((r) => r.no ? `${r.no}(${r.row}行)` : `${r.row}行`).join("、");
  return count > rows.length ? `${s} ほか ${count - rows.length} 行` : (s || "なし");
}
function stepTable(run) {
  const box = el("div");
  if (run.error) { box.append(el("p", "maint__msg", run.error)); return box; }
  const t = el("table", "table table--dense");
  const head = el("tr");
  for (const h of ["順番", "ルール", "ロットの値", "残った行の数", "残った行(管理番号・行)"]) head.append(el("th", "", h));
  t.append(el("thead"));
  t.tHead.append(head);
  const tb = el("tbody");
  for (const s of run.steps) {
    const tr = el("tr");
    for (const v of [s.order, s.rule, s.value, s.count, rowsText(s.rows, s.count)]) tr.append(el("td", "", String(v)));
    tb.append(tr);
  }
  t.append(tb);
  const wrap = el("div", "tablewrap");
  wrap.append(t);
  box.append(wrap, el("p", "", `使う行: ${rowsText(run.used, run.used.length)} → 結果: ${run.result}`));
  return box;
}

$("rTrialForm").addEventListener("submit", async (ev) => {
  ev.preventDefault();
  try {
    const d = await api.post("/api/maint/match_rules/trial", { group: R.name, lot: $("rLot").value, edited: R.work, inputs: trialInputs() });
    const out = [el("p", "muted maint__small", d.note || `${d.lot || "直接入力"}: ${d.cases.length} 件(表 ${d.table_rows} 行)`)];
    for (const c of d.cases) {
      const card = el("section", "card");
      const h = el("header");
      h.append(el("h4", "", c.label), el("span", c.changed ? "lv lv--warn" : "lv", c.changed ? "結果が変わる" : "同じ"));
      const two = el("div", "maint__two");
      for (const [lab, run] of [["保存済み", c.saved], ["編集中", c.edited]]) {
        const s = el("div", "pad");
        s.append(el("h5", "", lab), stepTable(run));
        two.append(s);
      }
      card.append(h, two);
      out.push(card);
    }
    $("rTrial").replaceChildren(...out);
  } catch (err) { toastError(err); }
});

async function regression() {
  const d = await api.post("/api/maint/match_rules/regression", { group: R.name, edited: R.work });
  const out = [el("p", d.changed_count ? "maint__msg" : "muted maint__small",
    `${d.checked} 件を確かめました。結果が変わるもの: ${d.changed_count} 件`)];
  if (d.changed.length) {
    const ul = el("ul", "maint__small");
    for (const c of d.changed) ul.append(el("li", "", `${c.lot} ${c.label}: ${c.before} → ${c.after}`));
    out.push(ul);
  }
  $("rReg").replaceChildren(...out);
  return d;
}
$("rCheck").addEventListener("click", async () => { try { await regression(); } catch (err) { toastError(err); } });
$("rSave").addEventListener("click", async () => {
  try {
    const reg = await regression();
    if (reg.changed_count && !$("rReason").value.trim()) {
      $("rReason").setAttribute("aria-invalid", "true");
      $("rReason").focus();
      toast("結果が変わるロットがあります。理由を書いてから保存してください", "warn");
      return;
    }
    const d = await withUnlock(() => api.post("/api/maint/match_rules", { group: R.name, edited: R.work, reason: $("rReason").value }));
    if (!d) return;
    $("rReason").removeAttribute("aria-invalid");
    $("rReason").value = "";
    R.state = d;
    pick(R.name);
    toast(`照合ルール「${R.name}」を保存しました(全 PC 共通)`, "ok");
  } catch (err) { toastError(err); }
});

async function open() {
  try {
    R.state = await api.get("/api/maint/match_rules");
    pick(R.name || R.state.groups[0].name);
  } catch (err) { toastError(err); }
}
const tab = document.querySelector('#mTabs .tab[data-tab="rules"]');
if (tab) tab.addEventListener("click", open);
// 歪公差チェックの「照合ルール … 順番 N」から開く
window.addEventListener("match-rules:open", async (ev) => {
  R.name = ev.detail.group;            // タブを開いたときの open() がこの群を選ぶ
  if (tab) tab.click(); else await open();
  setTimeout(() => {
    const idx = R.work ? R.work.rules.findIndex((r) => r.order === ev.detail.order) : -1;
    const tr = idx >= 0 ? $("rRules").children[idx] : null;
    if (tr) { tr.classList.add("is-hit"); tr.scrollIntoView({ block: "center" }); }
  }, 600);
});
if (!$("pRules").hidden) open();
