/*
  品質基準画面(VBA の UF品質基準)。読み取り専用。
  カードの中身・参照元・該当の有無はサーバ(ufkensa/kensa/quality_session.py)が決める。ここは並べるだけ。
*/

import { api } from "../api.js";
import { toast, toastError } from "../toast.js";

const $ = (id) => document.getElementById(id);
const LEVEL_KIND = { info: "info", warn: "warn", error: "error" };
const STATUS_LABEL = { ok: "", none: "該当なし・不要", error: "エラー" };

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text ?? "";
  return e;
}

/** pairs: [見出し, 値, 注意(サーバが決める)] */
function kv(dl, pairs) {
  dl.replaceChildren(...pairs.flatMap(([k, v, warn]) => {
    const dd = el("dd", "", v || "-");
    if (warn) dd.dataset.warn = "1";
    return [el("dt", "", k), dd];
  }));
}

/** 見出しの組(見出しと値を 1 つの枠に入れて横に並べる) */
function hkv(dl, pairs) {
  dl.replaceChildren(...pairs.map(([k, v, warn]) => {
    const box = el("div");
    const dd = el("dd", "", v || "-");
    if (warn) dd.dataset.warn = "1";
    box.append(el("dt", "", k), dd);
    return box;
  }));
}

function renderHeader(d) {
  const h = d.header || {};
  const which = $("qWhich");
  which.textContent = d.lot ? (d.current ? "検査中のロット" : "別のロット") : "";
  which.className = "lv " + (d.current ? "lv--ok" : "lv--info");
  $("qEmpty").hidden = !!d.lot;
  if (!d.lot) { $("qHeader").replaceChildren(); return; }
  if (!d.found) {
    hkv($("qHeader"), [["ロット", d.lot], ["結果", d.status === "not_found" ? "SIKALOTNOW にロットがありません" : "仕掛DB を読めませんでした"]]);
    return;
  }
  hkv($("qHeader"), [
    ["ロット", h.lot], ["オーダー", h.order], ["検査No", h.ken_no],
    ["材質・調質", [h.zai, h.tyu].filter(Boolean).join(" ")], ["用途", [h.youto, h.youto_name].filter(Boolean).join(" ")],
    ["製造寸法", h.size], ["オーダー寸法", h.order_size],
    ["納入先", h.nou], ["得意先", h.tori], ["送り先", h.okuri],
    ["表面処理G", h.grade], ["ﾌﾗｯﾄﾈｽ等級", h.flat], ["直角度等級", h.taik], ["ﾗﾃﾗﾙﾎﾞｰ等級", h.lateral],
    ["設備仕様" + (h.line ? `(${h.line})` : ""), h.equipment, !h.equipment_ok], ["技術指示", h.additional],
  ]);
}

function renderCard(c) {
  const card = el("section", "card qcard");
  card.dataset.status = c.status;
  card.dataset.key = c.key;
  const head = el("header");
  head.append(el("h3", "", c.title));
  if (STATUS_LABEL[c.status]) head.append(el("span", "lv " + (c.status === "error" ? "lv--ng" : "lv--warn"), STATUS_LABEL[c.status]));
  card.append(head);
  const pad = el("div", "pad");
  const src = el("p", "qcard__src");
  src.append(el("b", "", "参照元"));
  for (const s of c.sources.length ? c.sources : ["(参照元の記載なし)"]) src.append(el("span", "chip", s));
  if (c.updated) src.append(el("small", "muted", `更新日 ${c.updated}`));
  pad.append(src);
  if (c.rows.length) {
    const dl = el("dl", "qual__kv");
    kv(dl, c.rows);
    pad.append(dl);
  }
  if (c.table) {
    const wrap = el("div", "tablewrap qcard__table");
    const t = el("table", "table table--dense");
    const tr = el("tr");
    tr.append(...c.table.head.map((x) => el("th", "", x)));
    const thead = el("thead"); thead.append(tr);
    const tbody = el("tbody");
    tbody.append(...c.table.rows.map((r) => { const row = el("tr"); row.append(...r.map((x) => el("td", "", x))); return row; }));
    t.append(thead, tbody);
    wrap.append(t);
    pad.append(wrap);
  }
  if (c.note) pad.append(el("p", "qcard__note", c.note));
  for (const m of c.messages || []) if (m.text) pad.append(el("p", "qcard__msg", m.text));
  card.append(pad);
  return card;
}

function render(d) {
  renderHeader(d);
  for (const m of d.messages || []) if (m.text) toast(m.text, LEVEL_KIND[m.level] || "info");
  $("qKensaBox").hidden = !(d.kensa && d.kensa.length);
  if (d.kensa && d.kensa.length) hkv($("qKensa"), d.kensa);
  $("qCards").replaceChildren(...(d.cards || []).map(renderCard));
}

async function load(lot = "") {
  try {
    const d = await api.get(`/api/quality/lot${lot ? `?lot=${encodeURIComponent(lot)}` : ""}`);
    if (lot === "" && d.lot) $("qLot").value = d.lot;
    render(d);
  } catch (e) {
    toastError(e);
  }
}

$("qFind").addEventListener("submit", (e) => {
  e.preventDefault();
  const v = $("qLot").value.trim().toUpperCase();
  $("qLot").value = v;
  load(v);
});
$("qCurrent").addEventListener("click", () => { $("qLot").value = ""; load(""); });

load("");
