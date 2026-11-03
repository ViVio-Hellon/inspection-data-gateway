/*
  検査画面(A案「計測コックピット」)。
  判断はすべてサーバ(ufkensa/kensa/session.py)。ここは受け取った view を描き、操作を送るだけ。

  操作の返り値(Result): { view, messages, question, stop }
    messages … 知らせ(VBA の MsgBox の原文)。トーストで出す
    question … 聞くこと。確認窓で答えを聞き、answers に入れて同じ操作を呼び直す
    stop     … 止めた理由。状態は保ったまま(D10)

  測定器の入力を止めない(D6): 手入力欄(歪・ラテラル・外観・不良・検入高さ・BH)で Enter を押したら、
  入力先を測定値の欄(measIn)へ戻す。描き直しの前に焦点のあった欄は、描き直した後も同じ欄に戻す。
*/

import { api, tokenUrl } from "../api.js";
import { toast, toastError } from "../toast.js";
import { playMessage } from "../sound.js";
import { confirmBox, promptBox } from "../askbox.js";
import { packDialog } from "./pack_dialog.js";

const $ = (id) => document.getElementById(id);
const POS = ["S", "U", "C", "D", "E", "F"];
let V = null;
const pack = packDialog({ imageUrl: "/api/kensa/pack_image", send: (name, body) => op(name, body) });
const recent = [];

const LEVEL_KIND = { info: "info", warn: "warn", error: "error" };

async function op(name, body = {}, answers = {}) {
  let res;
  try {
    res = await api.post(`/api/kensa/${name}`, { ...body, answers });
  } catch (e) {
    toastError(e);
    return null;
  }
  for (const m of res.messages || []) {
    playMessage(m);                                               // 音(D29: 出来事の鍵。鳴らせなければ文言)
    if (m.text) toast(m.text, LEVEL_KIND[m.level] || "info");     // 音だけの知らせはトーストを出さない
  }
  if (res.view) render(res.view);
  if (res.question) {
    const q = res.question;
    const yes = await confirmBox(q.text, { ok: q.buttons === "okcancel" ? "OK" : "はい",
                                           cancel: q.buttons === "okcancel" ? "キャンセル" : "いいえ" });
    return op(name, body, { ...answers, [q.key]: !!yes });
  }
  if (res.stop) toast(res.stop.text, "error");
  if (res.view && res.view.report && res.view.report.report) toast(`帳票: ${res.view.report.report.html}`, "info");
  return res;
}

const refocus = () => $("measIn").focus();

function text(el, t) { el.textContent = t ?? ""; }

function el(tag, props = {}, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (k === "dataset") Object.assign(e.dataset, v);
    else if (k === "class") e.className = v;
    else if (v !== undefined && v !== null) e[k] = v;
  }
  e.append(...kids.filter((x) => x !== null && x !== undefined));
  return e;
}

function kv(dl, pairs) {
  dl.replaceChildren(...pairs.flatMap(([k, v]) => [el("dt", { textContent: k }), el("dd", { textContent: v ?? "" })]));
}

function cellTd(c, cls) {
  const td = document.createElement("td");
  if (cls) td.className = cls;
  if (!c || c.visible === false) { td.dataset.hidden = "1"; return td; }
  td.textContent = c.value ?? "";
  td.dataset.color = c.color || "";
  if (c.border) td.dataset.border = c.border;
  if (c.underline) td.dataset.underline = "1";
  if (c.name) {
    td.dataset.name = c.name;
    td.tabIndex = 0;
    td.title = c.name;
  }
  return td;
}

/** 手入力の欄(VBA の TextBox)。Enter で送って測定値の欄へ戻る。 */
function inputTd(c, name, { width = "5em", onFocus = null } = {}) {
  const td = document.createElement("td");
  td.className = "in";
  if (!c || c.visible === false) { td.dataset.hidden = "1"; return td; }
  td.dataset.color = c.color || "";
  const inp = el("input", { class: "cellin", value: c.value ?? "", autocomplete: "off", dataset: { name } });
  inp.style.width = width;
  if (c.enabled === false) inp.disabled = true;
  inp.addEventListener("keydown", async (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    await op("value", { name, value: inp.value });
    refocus();
  });
  if (onFocus) inp.addEventListener("focus", onFocus);
  td.append(inp);
  return td;
}

/** 選ぶ欄(VBA の ComboBox)。選んだら送って測定値の欄へ戻る。 */
function selectTd(c, name, choices) {
  const td = document.createElement("td");
  td.className = "in";
  if (!c || c.visible === false) { td.dataset.hidden = "1"; return td; }
  const sel = el("select", { class: "cellin", dataset: { name } });
  const list = [""].concat(choices || []);
  if (c.value && !list.includes(c.value)) list.push(c.value);
  sel.append(...list.map((x) => new Option(x || "-", x)));
  sel.value = c.value || "";
  sel.addEventListener("change", async () => { await op("value", { name, value: sel.value }); refocus(); });
  td.append(sel);
  return td;
}

function headRow(positions) {
  return el("tr", {}, el("th"), ...positions.map((p) => el("th", { textContent: p.visible ? p.label : "" })));
}

function rowOf(label, tds, cls) {
  const tr = el("tr", cls ? { class: cls } : {});
  tr.append(el("th", { class: "rowh", textContent: label }), ...tds);
  return tr;
}

function renderGrid(g) {
  const rows = [headRow(g.positions)];
  for (const r of g.rows) {
    if (!r.visible) continue;
    rows.push(rowOf(r.label, r.cells.map((c) => cellTd(c))));
  }
  if (g.diag.some((c) => c.visible !== false)) rows.push(rowOf(g.diag_label, g.diag.map((c) => cellTd({ ...c, name: "" }, "sub")), "sep"));
  // 直角度・ラテラルボーは VBA でも手で入れる TextBox
  if (g.diag_kou.some((c) => c.visible !== false)) {
    rows.push(rowOf(g.diag_kou_label, g.diag_kou.map((c) => inputTd(c, `Diag${c.pos}Kou`)), "sep"));
  }
  if (g.lateral.some((c) => c.visible !== false)) {
    rows.push(rowOf(g.lateral_label, g.lateral.map((c) => inputTd(c, `Lateral${c.pos}`)), "sep"));
  }
  if (g.ng.some((c) => c.visible !== false)) {
    const tr = rowOf("判定", g.ng.map((c) => {
      const td = cellTd({ ...c, name: "" }, "sub ng");
      if (c.visible !== false) { td.dataset.pos = c.pos; td.title = "ダブルクリックでまとめて削除"; }
      return td;
    }), "sep");
    rows.push(tr);
  }
  $("mgrid").replaceChildren(...rows);
}

function renderStrain(v) {
  const pos = v.grid.positions;
  const fr = v.strain_frame || {};
  const cb = v.cb_choices || {};
  const sc = v.strain.cb;
  const vis = (p) => (fr[p] ? fr[p].visible : true) && (pos.find((x) => x.key === p) || {}).visible !== false;
  const each = (fn) => POS.map((p) => (vis(p) ? fn(p) : el("td", { dataset: { hidden: "1" } })));
  const zLabel = POS.map((p) => (fr[p] && fr[p].lz) || "定盤").find((x) => x) || "定盤";
  const gLabel = POS.map((p) => (fr[p] && fr[p].lg) || "ｹﾞｰｼﾞ").find((x) => x) || "ｹﾞｰｼﾞ";
  const zc = Object.fromEntries(v.strain.z.map((c) => [c.pos, c]));
  const gc = Object.fromEntries(v.strain.g.map((c) => [c.pos, c]));
  $("strainGrid").replaceChildren(
    headRow(pos),
    rowOf(`${zLabel} 形状`, each((p) => selectTd(sc[`CB${p}1`], `CB${p}1`, (cb[p] || {})["1"]))),
    rowOf("上下・型", each((p) => selectTd(sc[`CB${p}2`], `CB${p}2`, (cb[p] || {})["2"]))),
    rowOf(`${zLabel} 値`, each((p) => inputTd(zc[p], `Z${p}`))),
    rowOf(`${gLabel} 形状`, each((p) => selectTd(sc[`CB${p}3`], `CB${p}3`, (cb[p] || {})["3"])), "sep"),
    rowOf("上下", each((p) => selectTd(sc[`CB${p}4`], `CB${p}4`, (cb[p] || {})["4"]))),
    rowOf(`${gLabel} 値`, each((p) => inputTd(gc[p], `G${p}`))),
  );
  text($("faceLv"), v.strain.face === "Now_B" ? "B面" : "A面");
  $("aorbBtn").textContent = v.strain.face === "Now_B" ? "いま B面(A面へ)" : "いま A面(B面へ)";
  text($("strainNote"), v.strain_note || "");
}

function fillList(sel, items, label) {
  const keep = new Set([...sel.selectedOptions].map((o) => o.value));
  const sig = JSON.stringify(items);
  if (sel.dataset.sig !== sig) {
    sel.replaceChildren(...items.map((it, i) => new Option(label(it), String(i))));
    sel.dataset.sig = sig;
    for (const o of sel.options) o.selected = keep.has(o.value);
  }
}

function renderVisual(v) {
  fillList($("visualList"), v.lists.visual, (x) => x || "(空白)");
  const caps = v.visual_caps || {};
  const vis = Object.fromEntries(v.visual.map((c) => [c.no, c]));
  const line = (label, base, key) => rowOf(label, POS.map((p, i) => {
    const n = base + i, c = vis[n] || {}, cap = caps[`${p}_${key}`] || {};
    if (c.visible === false) return el("td", { dataset: { hidden: "1" } });
    const td = el("td", { class: "lc", tabIndex: 0, textContent: cap.value || "", title: `LC${n}`, dataset: { lc: String(n), color: c.color || "" } });
    return td;
  }));
  $("visualGrid").replaceChildren(headRow(v.grid.positions), line("上面", 1, "UpVisual"), line("下面", 7, "UnderVisual"));
}

function renderBad(v) {
  fillList($("defectList"), v.lists.defect, (x) => `${x.name || "(空白)"}  ${x.code || ""}`);
  const bad = Object.fromEntries(v.bad.map((c) => [c.pos, c]));
  $("badGrid").replaceChildren(headRow(v.grid.positions),
    rowOf("枚数", POS.map((p) => inputTd(bad[p], `${p}Bad`, {
      width: "4em",
      // VBA の ?Bad_MouseDown(枚数を整えて水色に = 反映の書く先)
      onFocus: (e) => { if (!e.target.dataset.sent) { e.target.dataset.sent = "1"; op("value", { name: `${p}Bad`, value: e.target.value }); } },
    }))));
}

function renderKenH(v) {
  const box = $("yamaLb");
  box.replaceChildren(...(v.lists.yama || []).map((y) => {
    const b = el("button", { type: "button", class: "btn", textContent: y === "枚数" ? y : `${y}山` });
    b.setAttribute("aria-pressed", String(v.yama_lb === y));
    b.addEventListener("click", async () => { await op("yamalb", { value: y }); refocus(); });
    return b;
  }));
  const hl = v.ken_hl || {};
  const items = [];
  for (const c of v.ken_h) {
    if (c.visible === false) continue;
    const cap = c.no <= 6 ? ((hl[`HL${c.no}`] || {}).value || `${c.no}`) : "枚数";
    const unit = c.no <= 6 ? "mm" : "枚";
    const t = el("table", { class: "mgrid" }, el("tr", {}, el("th", { class: "rowh", textContent: cap }), inputTd(c, `KenHBox${c.no}`), el("th", { textContent: unit })));
    items.push(t);
  }
  if (v.ken_count_visible) {
    items.push(el("table", { class: "mgrid" }, el("tr", {}, el("th", { class: "rowh", textContent: "検査回数" }),
      inputTd({ value: v.ken_count, visible: true }, "KenCount"), el("th", { textContent: "回" }))));
  }
  $("kenhBoxes").replaceChildren(...items);
  text($("kenFormula"), [v.ken_h_label && `検入 ${v.ken_h_label}`, v.ken_formula].filter(Boolean).join("  "));
}

function renderBh(v) {
  const bh = Object.fromEntries(v.bh.map((c) => [c.pos, c]));
  $("bhGrid").replaceChildren(headRow(v.grid.positions.slice(0, 5)), rowOf("BH", ["S", "U", "C", "D", "E"].map((p) => inputTd(bh[p], `${p}BH`))));
  kv($("bhKv"), [["QF", v.bh_note.qf], ["LastHot", v.bh_note.last_hot]]);
}

function renderOrder(v) {
  const o = v.order;
  const sel = $("orderSel");
  const sig = JSON.stringify(o.list);
  if (sel.dataset.sig !== sig) {
    sel.replaceChildren(new Option("-", ""), ...o.list.map((x) => new Option(x, x)));
    sel.dataset.sig = sig;
  }
  sel.value = o.value || "";
  kv($("orderKv"), [["包装仕様No", o.no], ["梱包形態", o.form], ["タイトル", o.comment]]);
  const box = $("vcBox");
  const vc = v.vc;
  if (!vc) { box.replaceChildren(); return; }
  const side = (s) => (s.visible ? el("figure", {}, s.image ? el("img", { src: tokenUrl(`/api/kensa/picture/vc/${encodeURIComponent(s.image)}`), alt: s.label }) : null, el("figcaption", { textContent: s.label })) : null);
  box.replaceChildren(...[side(vc.a), side(vc.b), vc.nothing ? el("p", { class: "muted", textContent: vc.nothing }) : null].filter(Boolean));
}

// 包装仕様書(UF_PackNo)。全ヵ所・キャスクと同じ窓(pack_dialog.js)。閉じたら測定値の欄へ戻る
function renderPack(v) {
  const was = !!(V && V.pack);
  pack.render(v.pack);
  if (was && !v.pack) refocus();
}

// D28: ロットごとの「控え(第2)から読む」。決めるのはサーバ。ここは押された・入力欄の LotNo を送るだけ
function renderLotSource(v) {
  const s = v.lot_source || {};
  const btn = $("lotSrcBtn");
  btn.setAttribute("aria-pressed", String(!!s.on));
  btn.textContent = s.on ? "控え(第2)から読む: 入" : "控え(第2)から読む";
  $("lotSrcLv").hidden = !s.active;
  $("lotSrcLv").textContent = s.active ? `控えから読んでいます(${s.lot})` : "";
  const note = $("lotSrcNote");
  note.hidden = !s.on;
  note.textContent = !s.on ? "" : s.active
    ? `Lot ${s.lot} は、モード(いま ${s.mode})によらず控え(第2パス)の仕掛台帳から読みます。新規検査・別のロットの登録で外れます。`
    : `次に Lot登録 する ${s.lot} を控え(第2パス)から読みます。`;
}

function renderLeft(v) {
  renderLotSource(v);
  const h = v.header, t = v.tolerance;
  kv($("lotKv"), [["ライン", h.line + (h.line_no || "")], ["Lot", h.lot], ["検査No", h.ken_no], ["用途", h.youto],
                  ["材質/調質", [h.zai, h.tyu].filter(Boolean).join(" ")], ["コメント", h.comment]]);
  kv($("tolKv"), [["厚", t.atu && `${t.atu}  +${t.ap}/-${t.am}  (${t.amk}〜${t.apk})`],
                  ["巾", t.hab && `${t.hab}  +${t.hp}/-${t.hm}  (${t.hmk}〜${t.hpk})`],
                  ["丈", t.tak && `${t.tak}  +${t.tp}/-${t.tm}  (${t.tmk}〜${t.tpk})`],
                  ["対角", [t.taik, v.taik_text].filter(Boolean).join(" / ")], ["ラテラル", t.lateral_text || t.lateral_grade], ["フリーソリ", t.flatnes_text]]);
  text($("tolNote"), [t.additional, t.conductiv_text].filter(Boolean).join(" / "));
  const rs = v.buttons.resize;
  $("resizeLv").hidden = !rs;
  $("tolEdit").hidden = !rs;
  if (rs) {
    const f = [["ATU", "厚", t.atu], ["Ap", "厚+", t.ap], ["Am", "厚-", t.am], ["Hab", "巾", t.hab], ["Hp", "巾+", t.hp], ["Hm", "巾-", t.hm],
               ["Tak", "丈", t.tak], ["Tp", "丈+", t.tp], ["Tm", "丈-", t.tm], ["TAIK", "対角等級", t.taik], ["TaikText", "対角公差", v.taik_text],
               ["Lateral", "ラテ等級", t.lateral_grade], ["LateralText", "ラテ公差", t.lateral_text]];
    $("tolEdit").replaceChildren(...f.map(([name, label, val]) => {
      const lab = el("label", { textContent: label });
      const inp = el("input", { class: "input", value: val ?? "", dataset: { name }, autocomplete: "off" });
      inp.addEventListener("keydown", async (e) => { if (e.key === "Enter") { e.preventDefault(); await op("value", { name, value: inp.value }); refocus(); } });
      lab.append(inp);
      return lab;
    }));
  }
  const works = $("works");
  works.replaceChildren(...h.works.filter((w) => w.visible).map((w) => {
    const b = el("button", { type: "button", class: "btn", textContent: w.caption });
    b.setAttribute("aria-pressed", String(h.work === w.key));
    b.addEventListener("click", () => op("work", { work: w.key }));
    return b;
  }));
  $("yamaBox").hidden = !h.yama_visible;
  if (h.yama_visible && $("yamaSel").options.length === 0) {
    $("yamaSel").append(...["", "1", "2", "3", "4", "5", "6", "7"].map((y) => new Option(y || "-", y)));
  }
  $("yamaSel").value = h.yama || "";
  for (const [id, b] of [["atBtn", h.at_butt], ["inspeBtn", h.inspe_butt]]) {
    $(id).hidden = !b.visible; $(id).textContent = b.caption;
  }
  $("byousyaChk").checked = !!v.buttons.byousya;
  text($("atCount"), h.at_count);
  $("proofLv").textContent = h.proofread ? "校正済" : "未校正";
  $("proofLv").className = "lv " + (h.proofread ? "lv--ok" : "lv--ng");
  $("workerLv").hidden = !v.worker_needed;
  $("calibLink").hidden = !v.calibration_required;   // 校正_Open: 校正が要るときだけ校正画面への入口を出す
  $("workers").replaceChildren(...(v.workers || []).map((w) => el("li", { textContent: w })));
}

function renderRight(v) {
  text($("nextName"), v.next.label || v.next.name || "-");
  text($("nextHint"), [v.next.machin, v.next.auto ? "" : "手で選んだ欄"].filter(Boolean).join(" / "));
  const m = v.minmax;
  kv($("mmKv"), [["厚", m.amin && `${m.amin} 〜 ${m.amax}`], ["巾", m.hmin && `${m.hmin} 〜 ${m.hmax}`], ["丈", m.tmin && `${m.tmin} 〜 ${m.tmax}`]]);
  if (v.checks) {
    $("checks").replaceChildren(...v.checks.map((c) => el("li", { textContent: c.label, dataset: { ok: c.ok === null ? "" : (c.ok ? "1" : "0") } })));
  }
  $("blind").hidden = !v.blind;
  text($("blind"), v.blind);
  text($("workLv"), [v.header.work, v.header.yama && `${v.header.yama}山`].filter(Boolean).join(" "));
  const b = v.buttons;
  $("printBtn").textContent = b.print === "再印刷" ? "再印刷" : "印刷・保存";
  $("printBtn").dataset.printed = b.print === "再印刷" ? "1" : "";
  $("printResetBtn").hidden = !b.print_reset;
  $("abnoBtn").textContent = b.abno || "異常値on";
  $("izyouBtn").dataset.on = b.izyou ? "1" : "";
  $("resizeBtn").setAttribute("aria-pressed", String(!!b.resize));
  const nl = $("nameListBox");
  nl.hidden = !v.name_list;
  if (v.name_list) {
    const msg = v.name_list === 1 ? "登録検査員と測定者が一致しません。「組長」に連絡して対処の事。" : "梱包数と検入枚数が一致しません。日報を確認してください。";
    const a = el("a", { class: "btn btn--primary", href: `/meibo?mode=${v.name_list}`, textContent: "名簿・日報を開く" });
    a.addEventListener("click", async (e) => { e.preventDefault(); await op("namelist-ack"); location.href = a.href; });
    nl.replaceChildren(el("p", { textContent: msg }), a);
  }
}

// Min/Max の範囲(OPS1〜6 / OPE1〜6。VBA の OptionButton)
function renderRange(v) {
  const r = v.range;
  if (!r) return;
  for (const [id, cur] of [["opsSel", r.ops], ["opeSel", r.ope]]) {
    const sel = $(id);
    sel.replaceChildren(new Option("-", ""),
      ...r.labels.map((l, i) => (l ? new Option(l, String(i + 1)) : null)).filter(Boolean));
    sel.value = cur ? String(cur) : "";
  }
}

// 公差外の確認(UFConfi)。page 0 = 公差外(TB1 に入れて送信)、1 = 消去理由(List1)
function renderConfi(v) {
  const box = $("confiBox");
  const c = v.confi;
  box.hidden = !c;
  if (!c) return;
  const head = el("p", { textContent: c.page === 0 ? `検査ﾃﾞｰﾀに公差外有り！ 対象データ ${c.moto}(${c.posi}${c.n})`
                                                   : `再測定及びその他の理由 対象データ ${c.moto}(${c.posi}${c.n})` });
  const row = el("div", { class: "btn-row" });
  const name = c.page === 0 ? (c.list2 ? "UFConfi.List2" : "UFConfi.TB1") : "UFConfi.List1";
  row.append(...c.choices.map((ch) => {
    const b = el("button", { type: "button", class: "btn", textContent: ch });
    b.addEventListener("click", async () => {
      if (ch === "その他" || ch === "再測定") {
        const other = await promptBox("理由を入力してください", { value: "" });
        if (other === null) return;
        await op("value", { name: "UFConfi.その他", value: other });
      }
      await op("value", { name, value: ch });
      refocus();
    });
    return b;
  }));
  box.replaceChildren(head, row);
}

// 足した値が無い view(セッションの外で作ったもの)でも描けるように
const EMPTY = { lists: { visual: [], yama: [], defect: [] }, cb_choices: {}, order: { list: [] }, buttons: {}, visual_caps: {},
                ken_hl: {}, strain_frame: {}, bh_note: {}, pack: null, vc: null };

function render(v) {
  v = { ...EMPTY, ...v };
  renderPack(v);
  V = v;
  const act = document.activeElement;
  const keep = act && act.dataset && act.dataset.name && act.closest("#kensa") ? act.dataset.name : null;
  renderLeft(v);
  renderGrid(v.grid);
  renderStrain(v);
  renderVisual(v);
  renderBad(v);
  renderKenH(v);
  renderBh(v);
  renderOrder(v);
  renderRight(v);
  renderRange(v);
  renderConfi(v);
  if (keep) {                                   // 描き直しで焦点が外れないように
    const again = document.querySelector(`#kensa [data-name="${CSS.escape(keep)}"]`);
    if (again && again.tagName !== "TD") { again.dataset.sent = "1"; again.focus(); }
  }
}

function buildExtras() {
  const mm = $("mmKv").parentElement;
  const rg = el("div", { class: "btn-row" });
  for (const [id, label, prefix] of [["opsSel", "開始", "OPS"], ["opeSel", "終了", "OPE"]]) {
    const lab = el("label", { textContent: label });
    const sel = el("select", { class: "input", id });
    sel.addEventListener("change", () => { if (sel.value) op("value", { name: prefix + sel.value, value: "True" }); });
    lab.append(sel);
    rg.append(lab);
  }
  mm.append(rg);
  const confi = el("div", { id: "confiBox", class: "pad", hidden: true });
  $("checks").parentElement.before(confi);
}

function remember(line) {
  recent.unshift(`${new Date().toLocaleTimeString("ja-JP")} ${line}`);
  recent.length = Math.min(recent.length, 8);
  $("recent").replaceChildren(...recent.map((l) => el("li", { textContent: l })));
}

function onEnter(id, fn) {
  $(id).addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    const val = $(id).value;
    $(id).value = "";
    if (val) fn(val);
  });
}

const selected = (sel) => [...sel.selectedOptions].map((o) => Number(o.value));

function initTabs() {
  const tabs = [...document.querySelectorAll("#subtabs [data-tab]")];
  let cur = "strain";
  try { cur = localStorage.getItem("kensa.subtab") || cur; } catch (e) { /* 保存できなくても動く */ }
  const show = (name) => {
    for (const t of tabs) t.setAttribute("aria-selected", String(t.dataset.tab === name));
    for (const p of document.querySelectorAll(".subpane")) p.hidden = p.dataset.pane !== name;
    try { localStorage.setItem("kensa.subtab", name); } catch (e) { /* 同上 */ }
  };
  for (const t of tabs) t.addEventListener("click", () => { show(t.dataset.tab); refocus(); });
  show(tabs.some((t) => t.dataset.tab === cur) ? cur : "strain");
}

function init() {
  buildExtras();
  initTabs();
  onEnter("cardIn", (card) => op("worker", { card }));
  onEnter("lotIn", (lot) => op("lot", { lot }));
  $("lotSrcBtn").addEventListener("click", () => {
    const on = !(V && V.lot_source && V.lot_source.on);
    op("lot-source", { on, lot: $("lotIn").value });
  });
  onEnter("measIn", async (t) => { remember(t); await op("receive", { text: t }); refocus(); });
  onEnter("manualIn", async (t) => { remember(`手入力 ${t}`); await op("value", { name: "Manual", value: t }); refocus(); });
  $("yamaSel").addEventListener("change", () => op("yama", { yama: $("yamaSel").value }));
  $("atBtn").addEventListener("click", () => op("at"));
  $("inspeBtn").addEventListener("click", () => op("inspe"));
  $("byousyaChk").addEventListener("change", () => op("value", { name: "Byousya", value: $("byousyaChk").checked ? "True" : "False" }));
  $("mmBtn").addEventListener("click", () => op("minmax"));
  $("printBtn").addEventListener("click", async () => {
    const r = await op("print-check");
    if (r && !r.stop && r.view && r.view.print_ready) await op("print-commit");
  });
  $("printResetBtn").addEventListener("click", () => op("print-reset"));
  $("newBtn").addEventListener("click", () => op("new"));
  $("holdBtn").addEventListener("click", async () => { await op("hold"); refocus(); });
  $("abnoBtn").addEventListener("click", async () => { await op("value", { name: "Abno", value: "" }); refocus(); });
  $("izyouBtn").addEventListener("click", () => op("izyou"));
  $("resizeBtn").addEventListener("click", () => op("resize"));
  $("aorbBtn").addEventListener("click", async () => { await op("value", { name: "AorB", value: "" }); refocus(); });
  $("mgrid").addEventListener("click", (e) => {
    const td = e.target.closest("td[data-name]");
    if (td) { op("click", { name: td.dataset.name }); refocus(); }
  });
  $("mgrid").addEventListener("dblclick", async (e) => {
    const ng = e.target.closest("td.ng[data-pos]");
    if (ng) { await op("ng-delete", { pos: ng.dataset.pos }); return; }
    const td = e.target.closest("td[data-name]");
    if (!td || !td.textContent) return;
    op("dblclick", { name: td.dataset.name });
  });
  // 外観: 欠陥名を選んで LC を押す(LcColor)。ダブルクリックで白に戻して選択を外す(LC?_DblClick)
  $("visualGrid").addEventListener("click", async (e) => {
    const td = e.target.closest("td[data-lc]");
    if (td) { await op("visual", { lc: Number(td.dataset.lc), items: selected($("visualList")) }); refocus(); }
  });
  $("visualGrid").addEventListener("dblclick", async (e) => {
    const td = e.target.closest("td[data-lc]");
    if (!td) return;
    for (const o of $("visualList").options) o.selected = false;
    await op("visual-clear", { lc: Number(td.dataset.lc) });
  });
  $("reflectBtn").addEventListener("click", async () => {
    await op("defect-reflect", { items: selected($("defectList")) });
    refocus();
  });
  $("cleBtn").addEventListener("click", () => op("defect-clear"));
  $("orderSel").addEventListener("change", async () => { await op("order", { order: $("orderSel").value }); refocus(); });
  $("ordersBtn").addEventListener("click", () => op("orders"));
  api.get("/api/kensa/state").then((res) => { render(res.view); return op("start"); }).catch(toastError);
  refocus();
}

init();
