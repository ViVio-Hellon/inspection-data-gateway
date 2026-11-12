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
import { promptBox } from "../askbox.js";
import { LEVEL_KIND, screenOp } from "../screen_op.js";
import { fromServer, startReceiver } from "../recv.js";
import { alarmBox, collectPacking, isAbnormal, showJudgeImage } from "../judgebox.js";
import { packDialog } from "./pack_dialog.js";

const $ = (id) => document.getElementById(id);
const POS = ["S", "U", "C", "D", "E", "F"];
let V = null;
const pack = packDialog({ imageUrl: "/api/kensa/pack_image", send: (name, body) => op(name, body) });
const recent = [];

/** サーバの知らせ 1 件。2026-11-09: 異常値は大きな赤い窓(確認を押すまで)、合格・不合格・対角線は判定の画像も出す */
function showMessage(m) {
  playMessage(m);                                               // 音(D29: 出来事の鍵。鳴らせなければ文言)
  showJudgeImage(m);                                            // VBA の UFOK / UFNG / UF対角(画像が無ければ出さない)
  if (isAbnormal(m)) { alarmBox(m.text); return; }              // VBA の UF異常値発見
  if (collectPacking(m)) return;                                // 2026-11-11: 包装仕様の注意・梱包保護材は 1 つの窓に
  if (m.text) toast(m.text, LEVEL_KIND[m.level] || "info");     // 音だけの知らせはトーストを出さない
}

/* 操作を送る(screen_op.js)。知らせは showMessage、印刷のあとは帳票の案内 */
const op = screenOp("kensa", {
  render,
  showMessage,
  after: (name, res) => {
    if (name === "print-commit" || name === "print-check") announceReport(res.view);
    if (PREVIEW_OPS.has(name)) refreshDefectPreview();
  },
});

/* 不良タブの見本(2026-11-11): 印刷・保存と同じ 不良要因票(/api/kensa/defect_preview)。反映・ｸﾘｱのたびに描き直す */
const PREVIEW_OPS = new Set(["start", "lot", "lot-reload", "new", "defect-reflect", "defect-clear", "print-commit"]);
let previewSeq = 0;
async function refreshDefectPreview() {
  const frame = document.getElementById("defPrev"), empty = document.getElementById("defPrevEmpty");
  if (!frame) return;
  const seq = ++previewSeq;
  let html = "";
  try {
    const r = await fetch(tokenUrl("/api/kensa/defect_preview"), { cache: "no-store" });
    if (r.status === 200) html = await r.text();
  } catch (e) { /* 見本が出せなくても入力はできる */ }
  if (seq !== previewSeq) return;
  frame.hidden = !html;
  empty.hidden = !!html;
  if (!html) { frame.removeAttribute("srcdoc"); return; }
  frame.onload = fitDefectPreview;
  frame.srcdoc = html;
}
/** 帳票を箱の幅に合わせて縮める(縦は箱の中でスクロール) */
function fitDefectPreview() {
  const frame = document.getElementById("defPrev");
  const doc = frame && frame.contentDocument;
  if (!doc || !doc.body || frame.hidden) return;
  const sheet = doc.querySelector(".sheet") || doc.body;
  doc.body.style.zoom = "";
  const w = Math.max(sheet.scrollWidth, doc.documentElement.scrollWidth), avail = frame.clientWidth - 8;
  if (!w || !avail) return;
  doc.body.style.zoom = String(Math.min(1, avail / w));
  // 縮めた後もはみ出すとき(文字の最小の大きさなど)はもう一段縮める
  const over = doc.documentElement.scrollWidth - frame.clientWidth;
  if (over > 0) doc.body.style.zoom = String(Math.min(1, avail / w) * (frame.clientWidth - 8) / (frame.clientWidth + over));
}
window.addEventListener("resize", fitDefectPreview);

/** 印刷のあと 1 回だけ「帳票を開く」を知らせる(キャスクと同じく /api/…/report/<名前> で開く。サーバのパスは出さない)。 */
let lastReport = "";
function reportUrl(v) {
  const h = v && v.report && v.report.report && v.report.report.html;
  if (!h) return "";
  const name = String(h).split(/[\\/]/).pop();
  return tokenUrl(`/api/kensa/report/${encodeURIComponent(name)}`);
}
function announceReport(v) {
  const url = reportUrl(v);
  if (!url) return;
  const key = url + "|" + JSON.stringify((v.report && v.report.saved) || {}) + "|" + !!(v.report && v.report.reprint);
  if (key === lastReport) return;
  lastReport = key;
  toast(v.report.reprint ? "帳票を作り直しました" : "帳票を作りました", "ok",
        { label: "帳票を開く", onClick: () => window.open(url, "_blank", "noopener") });
}

const refocus = () => $("measIn")?.focus({ preventScroll: true });   // 画面を移ったあと(入れ替え式)でも落ちない

/**
 * 2026-11-10(現場の声): 検査員を登録して、この直の校正が済んでいなければ、そのまま校正画面へ移る。
 * VBA は 校正_Open で UF校正 を開いていた。`calibration_required` はサーバが決める
 * (校正をとばせる端末 = ufkensa:calibration_skip では立たない)。移ったら true。
 */
const CALIB_FROM_WORKER = "worker";
function toCalibrationAfterWorker(res) {
  const v = res && res.view;
  if (!v || !v.calibration_required || res.stop || res.question) return false;
  toast("この直の校正が済んでいません。校正画面へ移ります", "warn");
  setTimeout(() => { location.href = `/calibration?from=${CALIB_FROM_WORKER}`; }, 900);
  return true;
}

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

/* ---------- 測定者と時刻(WorkerTime。2026-11-08) ---------- */
const WHO_KEY = "kensa.showWho";
function showWho() { try { return localStorage.getItem(WHO_KEY) === "1"; } catch { return false; } }
function whoText(w) {
  if (!w) return "";
  const time = String(w.when || "").replace(/^\d{4}\/\d{1,2}\/\d{1,2}\s*/, "");     // 日付は省いて時刻だけ
  return [w.who, time].filter(Boolean).join(" ");
}

function cellTd(c, cls, wt) {
  const td = document.createElement("td");
  if (cls) td.className = cls;
  if (!c || c.visible === false) { td.dataset.hidden = "1"; return td; }
  td.textContent = c.value ?? "";
  const w = wt && c.name ? wt[c.name] : null;
  if (w && showWho() && c.value) {
    td.append(el("small", { class: "who", textContent: whoText(w) }));
  }
  td.dataset.color = c.color || "";
  if (c.border) td.dataset.border = c.border;
  if (c.underline) td.dataset.underline = "1";
  if (c.name) {
    td.dataset.name = c.name;
    td.tabIndex = 0;
    td.title = w ? `${c.name}  測定: ${w.who || "-"}  ${w.when || ""}${w.pake ? "  パケット " + w.pake : ""}` : c.name;
  }
  return td;
}

/** 手入力の欄(VBA の TextBox)。Enter で送って測定値の欄へ戻る。 */
function inputTd(c, name, { width = "5em", onFocus = null, sendOnChange = false } = {}) {
  const td = document.createElement("td");
  td.className = "in";
  if (!c || c.visible === false) { td.dataset.hidden = "1"; return td; }
  td.dataset.color = c.color || "";
  const inp = el("input", { class: "cellin", value: c.value ?? "", autocomplete: "off", dataset: { name } });
  inp.style.width = width;
  if (c.enabled === false) inp.disabled = true;
  let sentVal = inp.value;                      // 2 度送らないように、最後に送った値を覚える
  inp.addEventListener("keydown", async (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    sentVal = inp.value;
    await op("value", { name, value: inp.value });
    refocus();
  });
  // 2026-11-12(歪): Enter を押さずに次の欄へ移っても判定する(欄を離れた = change)。Enter で送った値は送らない
  if (sendOnChange) {
    inp.addEventListener("change", () => {
      if (inp.value === sentVal) return;
      sentVal = inp.value;
      op("value", { name, value: inp.value });
    });
  }
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

function headRow(positions, withTol = false) {
  const tr = el("tr", {}, el("th"), ...positions.map((p) => el("th", { textContent: p.visible ? p.label : "" })));
  if (withTol) tr.append(el("th", { class: "tolh", textContent: "公差(下限〜上限)" }));
  return tr;
}

/** 行の公差(表示だけ。判定はサーバのまま)。厚・巾・丈は頭の字で決める */
function tolOf(label, v) {
  const t = (v && v.tolerance) || {};
  const head = String(label || "").charAt(0);
  if (head === "厚" && t.atu) return `${t.amk}〜${t.apk}`;
  if ((head === "巾" || head === "幅") && t.hab) return `${t.hmk}〜${t.hpk}`;
  if (head === "丈" && t.tak) return `${t.tmk}〜${t.tpk}`;
  return "";
}
const tolTd = (txt) => el("td", { class: "tol", textContent: txt || "" });

function rowOf(label, tds, cls, tol) {
  const tr = el("tr", cls ? { class: cls } : {});
  tr.append(el("th", { class: "rowh", textContent: label }), ...tds);
  if (tol !== undefined) tr.append(tolTd(tol));
  return tr;
}

function renderGrid(g, v) {
  const t = (v && v.tolerance) || {};
  const rows = [headRow(g.positions, true)];
  let lastGroup = "", n = 0;
  for (const r of g.rows) {
    if (!r.visible) continue;
    // 厚・幅・丈ごとにまとまりの色と区切りを付け、公差はまとまりの最初の行にだけ出す
    const head = String(r.label || "").charAt(0);
    const group = head === "巾" ? "幅" : head;
    const first = group !== lastGroup;
    if (first) n += 1;
    lastGroup = group;
    const gk = { 厚: "atu", 幅: "hab", 丈: "tak", 対: "taik" }[group] || (n % 2 ? "a" : "b");
    const tr = rowOf(r.label, r.cells.map((c) => cellTd(c, "", v && v.worker_time)), `grp grp--${gk}${first ? " grp--first" : ""}`,
      first ? tolOf(r.label, v) : "");
    rows.push(tr);
  }
  if (g.diag.some((c) => c.visible !== false)) rows.push(rowOf(g.diag_label, g.diag.map((c) => cellTd({ ...c, name: "" }, "sub")), "sep",
    [t.taik, v && v.taik_text].filter(Boolean).join(" / ")));
  // 直角度・ラテラルボーは VBA でも手で入れる TextBox
  if (g.diag_kou.some((c) => c.visible !== false)) {
    rows.push(rowOf(g.diag_kou_label, g.diag_kou.map((c) => inputTd(c, `Diag${c.pos}Kou`)), "sep", ""));
  }
  if (g.lateral.some((c) => c.visible !== false)) {
    const tr = rowOf(g.lateral_label, g.lateral.map((c) => inputTd(c, `Lateral${c.pos}`)), "sep", t.lateral_text || t.lateral_grade || "");
    const hit = lateralHitText(v);
    if (hit) { tr.lastChild.append(el("small", { class: "lathit lathit--row", textContent: hit })); tr.classList.add("is-lathit"); }
    rows.push(tr);
  }
  if (g.ng.some((c) => c.visible !== false || NG_BADGE[String(c.value || "").trim()])) {
    const tr = rowOf("判定", g.ng.map((c) => {
      const td = cellTd({ ...c, name: "" }, "sub ng");
      if (c.visible !== false) { td.dataset.pos = c.pos; td.title = "ダブルクリックでまとめて削除"; }
      const badge = ngBadge(c.value, c.pos);
      if (badge && c.visible !== false) {
        td.replaceChildren(badge);
        td.dataset.flag = NG_BADGE[String(c.value).trim()];
        td.title = `${c.pos}: ${String(c.value).trim()}(ダブルクリックでまとめて削除${td.dataset.flag === "ng" ? "" : "。印を押すと外す"})`;
      }
      return td;
    }), "sep ngrow", "");
    rows.push(tr);
  }
  // 見出し(S・U・C…)にも同じ印を付ける。表の下まで目を移さなくても分かるように
  const flags = Object.fromEntries(g.ng.filter((c) => c.visible !== false).map((c) => [c.pos, String(c.value || "").trim()]));
  const ths = rows[0].querySelectorAll("th");
  g.positions.forEach((p, i) => {
    const b = p.visible ? ngBadge(flags[p.key], p.key) : null;
    if (b) { b.classList.add("ngflag--sm"); ths[i + 1].append(" ", b); }
  });
  $("mgrid").replaceChildren(...rows);
}

/* 2026-11-10: 公差を 寸法 / 歪 / 対角・ラテラル・フリー反り にまとめて出す(現場の声「歪の公差が見えない」)。表示だけ */
const STRAIN_TOL = [["巾", "ha"], ["丈", "ta"], ["中", "naka"], ["耳", "mimi"], ["MG", "mg"]];
function tolGroups(v) {
  const t = v.tolerance || {};
  const s = v.strain_tol || {};
  return [
    ["寸法", [["厚", t.atu && `${t.atu}  +${t.ap}/-${t.am}  (${t.amk}〜${t.apk})`],
              ["巾", t.hab && `${t.hab}  +${t.hp}/-${t.hm}  (${t.hmk}〜${t.hpk})`],
              ["丈", t.tak && `${t.tak}  +${t.tp}/-${t.tm}  (${t.tmk}〜${t.tpk})`]]],
    ["歪", STRAIN_TOL.map(([k, key]) => [k, s[key] || ""])],
    ["対角・ラテラル・フリー反り", [["対角", [t.taik, v.taik_text].filter(Boolean).join(" / ")],
                                  ["ラテラル", t.lateral_text || t.lateral_grade],
                                  ["フリー反り", t.flatnes_text || s.flatnes || ""]]],
  ];
}
/* 2026-11-12: ラテラルボー条件に当たった(VBA の Sumisei「ナメカワ」・LateralText.Tag="有")ことをずっと出す */
function lateralHitText(v) {
  const h = (v && v.lateral_hit) || {};
  if (!h.on) return "";
  const src = (h.source || []).join("・");
  return `ラテラルボー要: ${h.who || "技術指示"}${h.text ? ` 公差 ${h.text}` : ""}${src ? `(${src})` : ""}`;
}
function renderTolGroups(v) {
  const hit = lateralHitText(v);
  $("latHit").hidden = !hit;
  text($("latHit"), hit);
  $("tolKv").replaceChildren(...tolGroups(v).map(([title, pairs]) => {
    const dl = el("dl", { class: "kv-grid" });
    kv(dl, pairs.map(([k, x]) => [k, x || "-"]));
    for (const dd of dl.querySelectorAll("dd")) if (dd.textContent === "-") dd.classList.add("muted");
    return el("section", { class: "tol-group", dataset: { group: title } }, el("h4", { textContent: title }), dl);
  }));
}
function renderStrainTol(v) {
  const s = v.strain_tol || {};
  const has = STRAIN_TOL.some(([, key]) => s[key]);
  $("strainTol").replaceChildren(...(has ? [el("b", { textContent: "歪の公差" }),
    // 2026-11-11: 測定表の寸法公差と同じく、項目ごとに色の枠を付けて大きく出す(現場の声)
    ...STRAIN_TOL.map(([k, key]) => el("span", { class: "strain-tol__item" + (s[key] ? "" : " is-empty"), dataset: { k: key } },
                                                 el("small", { textContent: k }), el("span", { textContent: s[key] || "-" })))] : []));
}

/* 2026-11-10: NG? ラベル(NG / 許容 / 保留)をはっきり見せる(現場の声「保留・許容のフラグが見えない」) */
const NG_BADGE = { NG: "ng", 許容: "allow", 保留: "hold" };
function ngBadge(value, pos) {
  const v = String(value || "").trim();
  const kind = NG_BADGE[v];
  if (!kind) return null;
  const b = el("span", { class: `ngflag ngflag--${kind}`, textContent: v });
  if (pos && kind !== "ng") {          // 2026-11-11: 許容・保留は押すと外せる
    b.dataset.pos = pos;
    b.setAttribute("role", "button");
    b.title = `${pos}: ${v}(押すと${v}を外します。NG の値は消しません)`;
  }
  return b;
}

function renderStrain(v) {
  renderStrainTol(v);
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
    rowOf(`${zLabel} 値`, each((p) => inputTd(zc[p], `Z${p}`, { width: "6.5em", sendOnChange: true }))),
    rowOf(`${gLabel} 形状`, each((p) => selectTd(sc[`CB${p}3`], `CB${p}3`, (cb[p] || {})["3"])), "sep"),
    rowOf("上下", each((p) => selectTd(sc[`CB${p}4`], `CB${p}4`, (cb[p] || {})["4"]))),
    rowOf(`${gLabel} 値`, each((p) => inputTd(gc[p], `G${p}`, { width: "6.5em", sendOnChange: true }))),
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
  // 2026-11-11: VBA の ListBox と同じ 2 列(名前は左・コードは右)。<select> の中では列を揃えられないので、
  // 選べる一覧(role=listbox)で描く。名前が空(全角空白だけ)の行はコードだけを右に出す
  defectListbox($("defectList"), v.lists.defect || []);
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
  // 2026-11-11: 検査回数は少枚数のときだけ(厚板 H など他の作業では出さない。現場の声)
  if (v.ken_count_visible && v.header.work === "少枚数") {
    // 2026-11-10: 少枚数(厚板の検番 H)で枚数があるのに検査回数が無いと、測定はサーバで止まる
    // (input_router.ken_count_missing「検査回数を入れてください」)。同じ決まりで、欄を「必須」に見せる
    const thick = (v.header.works || []).find((w) => w.key === "厚板") || {};
    const sheets = (v.ken_h.find((c) => c.no === 7) || {}).value || "";
    const required = v.header.work === "少枚数" && thick.caption === "H" && sheets !== "" && !(Number(v.ken_count) >= 1);
    const td = inputTd({ value: v.ken_count, visible: true }, "KenCount");
    if (required) {
      td.dataset.required = "1";
      const inp = td.querySelector("input");
      if (inp) { inp.required = true; inp.placeholder = "必須"; inp.setAttribute("aria-invalid", "true"); }
    }
    items.push(el("table", { class: "mgrid" }, el("tr", {},
      el("th", { class: "rowh", textContent: required ? "検査回数(必須)" : "検査回数" }),
      td, el("th", { textContent: "回" }))));
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
  renderVcRail(v);
  if (!vc) { box.replaceChildren(); return; }
  const side = (s) => (s.visible ? el("figure", {}, s.image ? el("img", { src: vcUrl(s.image), alt: s.label }) : null, el("figcaption", { textContent: s.label })) : null);
  box.replaceChildren(...[side(vc.a), side(vc.b), vc.nothing ? el("p", { class: "muted", textContent: vc.nothing }) : null].filter(Boolean));
}

const vcUrl = (name) => tokenUrl(`/api/kensa/picture/vc/${encodeURIComponent(name)}`);

// 2026-11-10: VC は測定中ずっと右の列に出す(公差と同じく、見て確かめる)。押すと大きく
function renderVcRail(v) {
  const box = $("vcRailImgs");
  const vc = v.vc;
  const hasLot = !!(v.header && v.header.lot);
  $("vcRail").hidden = !hasLot;
  if (!hasLot) { box.replaceChildren(); box.dataset.sig = ""; return; }
  const sides = vc ? [vc.a, vc.b].filter((s) => s && (s.visible || s.label)) : [];
  const sig = JSON.stringify([vc && vc.a, vc && vc.b, vc && vc.nothing]);
  if (box.dataset.sig === sig) return;          // 毎回描き直すと画像がちらつく
  box.dataset.sig = sig;
  const none = !vc || (!sides.length && !vc.nothing);
  $("vcRail").dataset.none = none ? "1" : "";
  $("vcRailHint").hidden = none || !sides.some((s) => s.visible && s.image);
  if (none) { box.replaceChildren(el("p", { class: "vcrail__none", textContent: "VC なし" })); return; }
  const fig = (s) => {
    const ok = s.visible && s.image;
    const f = el("figure", { class: "vcrail__fig" + (ok ? "" : " is-missing") },
      el("figcaption", { textContent: s.label || "VC" }),
      ok ? el("img", { src: vcUrl(s.image), alt: s.label }) : el("small", { textContent: "画像なし" }));
    if (ok) {
      f.tabIndex = 0;
      f.title = "押すと大きく表示";
      f.addEventListener("click", () => openVc(vc));
    }
    return f;
  };
  box.replaceChildren(...sides.map(fig), ...(vc.nothing ? [el("p", { class: "vcrail__why", textContent: vc.nothing })] : []));
  box.dataset.count = String(sides.length);
}

function openVc(vc) {
  const items = [vc.a, vc.b].filter((s) => s && s.visible && s.image);
  $("kVcBody").replaceChildren(...items.map((s) => el("figure", {}, el("figcaption", { textContent: s.label }), el("img", { src: vcUrl(s.image), alt: s.label }))));
  $("kVcBody").dataset.count = String(items.length);
  if (!$("kVc").open) $("kVc").showModal();
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
  // 決まった値(登録済みのロット)を大きく。入力欄は「別のロット」用に控えめ
  const hasLot = !!h.lot;
  $("lotReg").hidden = !hasLot;
  $("lotEmpty").hidden = hasLot;
  $("lotInLabel").textContent = hasLot ? "別のロット" : "LotNo";
  text($("lotRegNo"), h.lot || "");
  // 2026-11-09: 製品サイズ(厚 × 巾 × 丈)を大きく出す(現場の声)
  text($("lotSize"), hasLot && (t.atu || t.hab || t.tak) ? `${t.atu || "-"} × ${t.hab || "-"} × ${t.tak || "-"}` : "");
  kv($("lotRegKv"), [["検査No", h.ken_no], ["用途", h.youto], ["材質/調質", [h.zai, h.tyu].filter(Boolean).join(" ")]]);
  kv($("lotKv"), [["ライン", h.line + (h.line_no || "")], ["コメント", h.comment]]);
  renderTolGroups(v);
  text($("tolNote"), [t.additional, t.conductiv_text].filter(Boolean).join(" / "));
  renderResize(v);
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
  // 登録済みの検査員(サーバの "T1: 名前(日付)")を大きく出す
  const ws = v.workers || [];
  $("workers").replaceChildren(...ws.map((w) => {
    const m = /^(T\d):\s*(.*?)(?:\(([^()]*)\))?$/.exec(w) || [null, "", w, ""];
    // この直に登録が無い(worker_needed)ときは、残っているのは前の直の登録
    const stale = !!v.worker_needed;
    return el("li", stale ? { class: "is-stale" } : {}, el("span", { class: "reg-badge" + (stale ? " reg-badge--stale" : ""),
      textContent: stale ? `${m[1] || ""} 前の直の登録` : (m[1] ? `${m[1]} 登録済` : "登録済") }),
      el("b", { class: "reg-workers__name", textContent: m[2] }), m[3] ? el("small", { class: "muted", textContent: m[3] }) : null);
  }));
  $("workerOk").hidden = !ws.length || !!v.worker_needed;
  $("workerOk").textContent = ws.length ? `登録済 ${ws.length}名` : "";
  $("cardIn").placeholder = ws.length ? "もう1人追加するとき(社員カード)" : "社員カード(13桁)";
}

// ｻｲｽﾞ修正(2026-11-10): 右の列では欄が切れて入れられなかったので、大きな窓に 厚・巾・丈 × 値/+/- の表で出す。
// 欄は一度だけ作り、描き直しでは値だけ入れ替える(打っている欄の焦点と文字を消さない)
let resizeWas = null;            // 最初の描画(読み直し)では開かない。「はい」で入った時だけ開く
function renderResize(v) {
  const t = v.tolerance, rs = !!v.buttons.resize;
  $("resizeLv").hidden = !rs;
  $("resizeOpen").hidden = !rs;
  const vals = { ATU: t.atu, Ap: t.ap, Am: t.am, Hab: t.hab, Hp: t.hp, Hm: t.hm, Tak: t.tak, Tp: t.tp, Tm: t.tm,
                 TAIK: t.taik, TaikText: v.taik_text, Lateral: t.lateral_grade, LateralText: t.lateral_text };
  const box = $("tolEdit");
  if (rs && !box.dataset.built) {
    box.dataset.built = "1";
    const cell = (name, label) => {
      const inp = el("input", { class: "input", dataset: { name }, autocomplete: "off", "aria-label": label });
      inp.addEventListener("keydown", async (e) => {
        if (e.key !== "Enter") return;
        e.preventDefault();
        await op("value", { name, value: inp.value });
        const all = [...box.querySelectorAll("input")];
        const nx = all[all.indexOf(inp) + 1];
        if (nx && $("kResize").open) { nx.focus(); nx.select(); } else if (!$("kResize").open) refocus();
      });
      return inp;
    };
    const head = ["", "値", "＋(上)", "－(下)"].map((x) => el("b", { class: "tol-edit__h", textContent: x }));
    const rows = [["厚", "ATU", "Ap", "Am"], ["巾", "Hab", "Hp", "Hm"], ["丈", "Tak", "Tp", "Tm"]].flatMap(([lab, a, p, m]) =>
      [el("b", { class: "tol-edit__k", textContent: lab }), cell(a, lab), cell(p, lab + "+"), cell(m, lab + "-")]);
    const extra = [["対角", "TAIK", "TaikText"], ["ラテラル", "Lateral", "LateralText"]].map(([lab, g, txt]) =>
      el("div", { class: "tol-edit__row" }, el("b", { class: "tol-edit__k", textContent: lab }),
        el("label", { textContent: "等級" }, cell(g, lab + "等級")), el("label", { textContent: "公差" }, cell(txt, lab + "公差"))));
    box.replaceChildren(el("div", { class: "tol-edit__grid" }, ...head, ...rows), ...extra);
  }
  for (const inp of box.querySelectorAll("input")) {
    if (inp !== document.activeElement) inp.value = vals[inp.dataset.name] ?? "";
  }
  if (rs && resizeWas === false && !$("kResize").open) openResize();   // 「はい」で入ったら開く
  if (!rs && $("kResize").open) $("kResize").close();
  resizeWas = rs;
}
function openResize() {
  const d = $("kResize");
  if (!d.open) d.showModal();
  const first = $("tolEdit").querySelector("input");
  if (first) { first.focus(); first.select(); }
}

/* ---------- ロット情報の読み直し・見張り(2026-11-07) ---------- */
const WATCH_MS = 60000;            // 1 分ごとに、仕掛・引当・オーダー・品質基準が変わっていないか見る
async function reloadLot() {
  $("lotReloadBtn").disabled = true;
  try {
    const r = await op("lot-reload");
    if (r && !r.stop) $("lotWatchBar").hidden = true;
  } finally { $("lotReloadBtn").disabled = false; }
}
async function watchLot() {
  if (document.hidden) return;
  let w;
  try { w = await api.get("/api/kensa/lot-watch"); } catch { return; }   // 見張りの失敗は黙って次へ
  const bar = $("lotWatchBar");
  if (!w.changed || !w.changed.length) { bar.hidden = true; return; }
  $("lotWatchText").textContent = `ロット ${w.lot} の情報が更新されています(${w.changed.join("・")})`;
  bar.hidden = false;
  if (w.new) toast(`ロット ${w.lot} の情報が更新されています(${w.changed.join("・")})。「読み直す」で取り込めます`, "warn");
}

/* ---------- 検査員の追加・ロットの読み直し(上の帯の 1・3 を押す) ---------- */
// 2026-11-12: 検査員とロットは別の窓(別の入力欄)。前に打ちかけた文字は開くとき・閉じるときに消す
function clearAddInputs() { $("cardIn").value = ""; $("lotIn").value = ""; }
function openAdd(id) {
  if (phase === "measure") {
    const d = $("kAdd");
    const card = id === "cardIn";
    $("cardBox").hidden = !card;
    $("lotBox").hidden = card;
    $("lotSrcRow").hidden = card;
    text($("kAddTitle"), card ? "検査員の追加" : "ロットの読み直し");
    text($("kAddHint"), card ? "社員カード(13桁)を読み取ると自動で受け付けます。"
      : "LotNo(7桁)を読み取ると自動で受け付けます。同じ LotNo を入れると読み直します。");
    if (d.open) d.close();
    clearAddInputs();
    d.showModal();
  }
  setTimeout(() => $(id).focus(), 0);
}

/* ---------- 選ぶまで(案A)/ 測定(案B)の切り替え ---------- */
let phase = "";
function setPhase(p) {
  if (p === phase) return;
  phase = p;
  $("kensa").dataset.phase = p;
  const measure = p === "measure";
  // 作業の選択肢: 選ぶまでは中央の大きな枠、測定中は上の行
  (measure ? $("kbarWorks") : $("chooser")).append($("workPad"));
  // 公差の詳しい欄(対角・ラテラル・サイズ修正)は測定中は右の列へ。測定表の行にも出す
  (measure ? $("rpaneTol") : $("kside")).append($("tolCard"));
  // 2026-11-09: 受信機の状態(入らない理由)は選ぶまでも見えるように左の列へ。測定中は右の列(次に測るの下)
  if (measure) $("nextBox").after($("recvBox")); else $("kside").prepend($("recvBox"));
  if (measure) setTimeout(refocus, 0);           // 2026-11-11: 測定に入ったら受け取る欄に焦点を置く
  // カードの追加・別のロットの入力は、測定中は右の列の下へ
  const add = $("kAddPad");
  if (measure) add.append($("cardBox"), $("lotBox"), $("lotSrcRow"));
  else {
    if ($("kAdd").open) $("kAdd").close();
    for (const id of ["cardBox", "lotBox", "lotSrcRow"]) $(id).hidden = false;
    $("workers").after($("cardBox"));
    $("lotEmpty").after($("lotBox"));
    $("lotBox").after($("lotSrcRow"));
  }
}

function renderPhase(v) {
  const h = v.header || {};
  const ws = v.workers || [];
  const workerOk = ws.length > 0 && !v.worker_needed;
  const calibOk = !!h.proofread;
  const lotOk = !!h.lot;
  const workOk = !!h.work;
  setPhase(lotOk && workOk ? "measure" : "select");
  const checks = v.checks || [];
  const checkOk = checks.length > 0 && checks.every((c) => c.ok);
  const printed = v.buttons && v.buttons.print === "再印刷";
  const steps = [
    ["検査員", workerOk, workerOk ? ws.map((w) => w.replace(/^T\d:\s*/, "").replace(/\(.*\)$/, "")).join("・") : "カードを読む"],
    ["校正", calibOk, calibOk ? "済" : "未校正"],
    ["ロット", lotOk, lotOk ? h.lot : "LotNo を入れる"],
    ["作業選択", workOk, workOk ? [h.work, h.yama && `${h.yama}山`].filter(Boolean).join(" ") : "選ぶ"],
    ["測定", null, workOk ? "次: " + ((v.next && (v.next.label || v.next.name)) || "-") : ""],
    ["歪・外観", null, ""],
    ["印刷前チェック", checks.length ? checkOk : null, checks.length ? (checkOk ? "OK" : "未了あり") : ""],
    ["印刷", printed ? true : null, printed ? "済" : ""],
  ];
  // いまの段 = 最初の済んでいない段(1〜4)。全部済めば「測定」
  let cur = steps.findIndex((x, i) => i < 4 && !x[1]);
  if (cur < 0) cur = checkOk ? (printed ? 7 : 7) : 4;
  $("ksteps").replaceChildren(...steps.map(([name, ok, sub], i) => {
    const st = ok === true ? "done" : (i === cur ? "cur" : (ok === false && i < cur ? "ng" : ""));
    const li = el("li", { class: "kstep" + (st ? " is-" + st : "") },
      el("b", { textContent: `${i + 1} ${name}` }), el("small", { textContent: sub || "" }));
    if (i === 0 || i === 2) {                    // 検査員・ロットの段を押すと、その入力を出す
      li.classList.add("is-click");
      li.title = i === 0 ? "押すと検査員のカードを追加できます" : "押すと LotNo を入れ直せます(同じ番号で読み直し)";
      li.addEventListener("click", () => openAdd(i === 0 ? "cardIn" : "lotIn"));
    }
    if (i === 1 && !calibOk) {
      const a = el("a", { href: "/calibration", textContent: "校正へ" });
      li.append(a);
    }
    return li;
  }));
  // 選ぶまでの中央: 何をすればよいか
  const wait = !lotOk ? (workerOk ? "LotNo を入れてください(左の「ロット」)" : "検査員のカードを読み、LotNo を入れてください")
    : "";
  $("kwait").hidden = !!(lotOk);
  text($("kwaitText"), wait);
  $("chooser").hidden = !(lotOk && !workOk);
  // 測定中の上の1行
  const info = $("kbarInfo");
  const chip = (cls, t) => el("span", { class: "lv " + cls, textContent: t });
  // 新規検査: 測定中は上の行、選ぶまで(印刷のあとを含む)は作業選択の見出しに出す
  const newHome = phase === "measure" ? $("kbarOps") : $("chooserOps");
  if ($("newBtn").parentElement !== newHome) newHome.append($("newBtn"));
  // 上の帯の「利用者」: 検査員を登録したらその名前を出す(ログインの名前が無いとき)
  const rb = document.getElementById("rb-user");
  if (rb && (rb.dataset.state === "unset" || rb.dataset.worker === "1")) {
    const names = workerOk ? ws.map((w) => w.replace(/^T\d:\s*/, "").replace(/\(.*\)$/, "")).join("・") : "";
    if (names) { rb.textContent = names; rb.dataset.worker = "1"; delete rb.dataset.state; }
    else if (rb.dataset.worker === "1") { rb.textContent = "未ログイン"; rb.dataset.state = "unset"; delete rb.dataset.worker; }
  }
  info.replaceChildren(
    chip(workerOk ? "lv--ok" : "lv--ng", workerOk ? "検査員" : "検査員登録なし"),
    el("span", { class: "kbar__names", textContent: ws.map((w) => w.replace(/^T\d:\s*/, "").replace(/\(.*\)$/, "")).join("・") }),
    chip(calibOk ? "lv--ok" : "lv--ng", calibOk ? "校正済" : "未校正"),
    el("b", { class: "kbar__lot", textContent: h.lot || "" }),
    // 2026-11-09: 測定中も製品サイズ(厚 × 巾 × 丈)を見える所に
    el("b", { class: "kbar__size", textContent: (v.tolerance && (v.tolerance.atu || v.tolerance.hab || v.tolerance.tak))
      ? `${v.tolerance.atu || "-"} × ${v.tolerance.hab || "-"} × ${v.tolerance.tak || "-"}` : "" }),
    el("span", { textContent: [h.ken_no && `検査No ${h.ken_no}`, h.youto && `用途 ${h.youto}`, [h.zai, h.tyu].filter(Boolean).join(" ")].filter(Boolean).join("　") }),
  );
}

function renderRight(v) {
  text($("nextName"), v.next.label || v.next.name || "-");
  text($("nextHint"), [v.next.machin, v.next.auto ? "" : "手で選んだ欄"].filter(Boolean).join(" / "));
  const m = v.minmax;
  kv($("mmKv"), [["厚", m.amin && `${m.amin} 〜 ${m.amax}`], ["巾", m.hmin && `${m.hmin} 〜 ${m.hmax}`], ["丈", m.tmin && `${m.tmin} 〜 ${m.tmax}`]]);
  // 2026-11-10: 位置ごとの判定の印(許容・保留の位置は Min/Max に入らない = VBA の BackColor_Check)
  const flagged = (v.grid.ng || []).filter((c) => c.visible !== false && NG_BADGE[String(c.value || "").trim()]);
  $("mmFlags").replaceChildren(...(flagged.length
    ? [el("small", { class: "muted", textContent: "位置の判定: " }),
       ...flagged.map((c) => el("span", { class: "mmflag" }, el("b", { textContent: c.pos }), ngBadge(c.value)))]
    : [el("small", { class: "muted", textContent: "位置の判定: 印なし(NG・許容・保留はありません)" })]));
  if (v.checks) {
    $("checks").replaceChildren(...v.checks.map((c) => el("li", { textContent: c.label, dataset: { ok: c.ok === null ? "" : (c.ok ? "1" : "0") } })));
    $("chkDot").hidden = !v.checks.some((c) => c.ok === false);
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
  // F4(2026-11-10): 描き直しのたびに作り直すと、開いた選び欄が閉じ・選んでいる途中で戻る。
  // 選択肢が変わったときだけ作り直し、選んでいる最中(焦点がある)は触らない
  const sig = JSON.stringify(r.labels);
  for (const [id, cur] of [["opsSel", r.ops], ["opeSel", r.ope]]) {
    const sel = $(id);
    if (document.activeElement === sel) continue;
    if (sel.dataset.sig !== sig) {
      sel.replaceChildren(new Option("-", ""),
        ...r.labels.map((l, i) => (l ? new Option(l, String(i + 1)) : null)).filter(Boolean));
      sel.dataset.sig = sig;
    }
    const want = cur ? String(cur) : "";
    if (sel.value !== want) sel.value = want;
  }
}

// 公差外の確認(UFConfi)。page 0 = 公差外(TB1 に入れて送信)、1 = 消去理由(List1)
// 2026-11-10: 選ぶまで次へ進めない大きな窓(showModal。Esc・外側では閉じない。サーバも測定を受け付けない)
let confiSig = "";
function renderConfi(v) {
  const dlg = $("confiDlg");
  const box = $("confiBox");
  const c = v.confi;
  if (!c) {
    confiSig = "";
    if (dlg.open) { dlg.close(); refocus(); }
    return;
  }
  const sig = JSON.stringify(c);
  if (sig === confiSig && dlg.open) return;            // 描き直しで選んだボタンが消えないように
  confiSig = sig;
  $("confiTitle").textContent = c.page === 0 ? "検査ﾃﾞｰﾀに公差外有り！" : (c.list2 ? "再測定の理由" : "再測定及びその他の理由");
  const head = el("p", { class: "confi__target", textContent: `対象データ ${c.moto}(${c.posi}${c.n})` });
  const lead = el("p", { class: "confi__lead", textContent: "理由を選ぶまで次の測定はできません" });
  const row = el("div", { class: "confi__choices" });
  const note = el("p", { class: "muted confi__note" });
  note.append("選択肢は ", el("a", { href: "/maint?tab=master", target: "_blank", rel: "noopener", textContent: "保守 → マスタ管理 → 理由選択" }), " で増やせます");
  // 公差外のページ(TB1)は 1 つ押せば送る(理由の文字は要らない。VBA の page0)。
  // 理由のページ(List1)と再測定の理由(List2)は、VBA と同じく複数選んでから「決定」(2026-11-08 に直した)
  if (c.page === 0 && !c.list2) {
    row.append(...c.choices.map((ch) => {
      const b = el("button", { type: "button", class: "btn", textContent: ch });
      b.addEventListener("click", async () => { await op("value", { name: "UFConfi.TB1", value: ch }); refocus(); });
      return b;
    }));
    box.replaceChildren(head, lead, row);
  } else {
    const name = c.list2 ? "UFConfi.List2" : "UFConfi.List1";
    const picked = new Set();
    row.append(...c.choices.map((ch) => {
      const b = el("button", { type: "button", class: "btn", textContent: ch });
      b.setAttribute("aria-pressed", "false");
      b.addEventListener("click", () => {
        if (picked.has(ch)) picked.delete(ch); else picked.add(ch);
        b.setAttribute("aria-pressed", String(picked.has(ch)));
      });
      return b;
    }));
    const ok = el("button", { type: "button", class: "btn btn--commit confi__ok", textContent: "決定" });
    ok.addEventListener("click", async () => {
      const list = c.choices.filter((ch) => picked.has(ch));
      // 「その他」(理由のページは「再測定」も)を選んだら、理由の文字を聞く(VBA: 理由を入力してください)
      if (picked.has("その他") || (c.page === 1 && picked.has("再測定"))) {
        const other = await promptBox("理由を入力してください", { value: "" });
        if (other === null) return;
        await op("value", { name: "UFConfi.その他", value: other });
      }
      await op("value", { name, value: list.join(",") });   // 何も選ばなければサーバが「選択してください」
      refocus();
    });
    box.replaceChildren(head, lead, row, el("div", { class: "btn-row confi__foot" }, ok));
  }
  box.append(note);
  if (!dlg.open) dlg.showModal();
  dlg.focus();                                         // 最初のボタンに焦点を置かない(測定器の Enter で押されないように)
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
  renderGrid(v.grid, v);
  renderPhase(v);
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
  const dlg = el("dialog", { id: "confiDlg", class: "kpop kpop--confi" });
  dlg.setAttribute("aria-labelledby", "confiTitle");
  dlg.append(el("header", {}, el("h3", { id: "confiTitle", textContent: "確認" })),
             el("div", { id: "confiBox", class: "pad" }));
  dlg.tabIndex = -1;
  dlg.addEventListener("cancel", (e) => e.preventDefault());   // Esc では閉じない(選ぶまで進めない)
  // 測定器(キーボードと同じ)の Enter で、焦点のあるボタンが勝手に押されないように。選ぶのはマウス・タッチだけ
  dlg.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      if (e.key === "Enter") toast("理由を選んでから測ってください(ボタンを押して選びます)", "warn");
    }
  });
  document.body.append(dlg);
}

function remember(line) {
  recent.unshift(`${new Date().toLocaleTimeString("ja-JP")} ${line}`);
  recent.length = Math.min(recent.length, 8);
  $("recent").replaceChildren(...recent.map((l) => el("li", { textContent: l })));
}

function onEnter(id, fn, autoLen = 0) {
  const box = $(id);
  if (!box) return;
  const send = () => {
    const val = box.value;
    box.value = "";
    if (val) fn(val);
  };
  box.addEventListener("keydown", (e) => {
    // VBA は Enter でも Tab でも受けた(U-WAVEPAK などの区切りが Tab の現場がある)。空の Tab は普通に次へ移る
    if (e.key !== "Enter" && !(e.key === "Tab" && box.value)) return;
    e.preventDefault();
    send();
  });
  // VBA と同じく、決まった桁数まで打った(読み取った)ら Enter を待たずに受け付ける
  // (UF1.Worker_Change: 13 桁 / Lot_Change: 7 桁)
  if (autoLen) box.addEventListener("input", () => { if (box.value.trim().length === autoLen) send(); });
}

const selected = (sel) => (sel.tagName === "SELECT" ? [...sel.selectedOptions]
  : [...sel.querySelectorAll('[role="option"][aria-selected="true"]')]).map((o) => Number(o.value ?? o.dataset.value));

/** 不良要因の一覧(2 列)。押すと選ぶ/外す(1 つだけ)。↑↓ で動き、Space で選ぶ。描き直しでも選択は残す */
function defectListbox(box, items) {
  const sig = JSON.stringify(items);
  if (box.dataset.sig === sig) return;
  const keep = new Set(selected(box));
  box.dataset.sig = sig;
  box.replaceChildren(...items.map((x, i) => {
    const name = String(x.name || "").trim(), code = String(x.code || "").trim();
    return el("div", { role: "option", class: "deflist__row", dataset: { value: String(i) }, ariaSelected: String(keep.has(i)) },
      el("span", { class: "deflist__name", textContent: name }), el("span", { class: "deflist__code", textContent: code }));
  }));
  for (const o of box.children) o.setAttribute("aria-selected", String(keep.has(Number(o.dataset.value))));
  if (box.dataset.wired) return;
  box.dataset.wired = "1";
  // 2026-11-11: 1 つだけ選ぶ(サーバは「不良要因は1つずつ選んでから枚数を入れてください」)。別の行を押すと入れ替え、
  // 選んだ行をもう一度押すと外す
  const toggle = (o) => {
    if (!o) return;
    const on = o.getAttribute("aria-selected") !== "true";
    for (const r of box.children) r.setAttribute("aria-selected", String(on && r === o));
  };
  box.addEventListener("click", (e) => {
    const o = e.target.closest('[role="option"]');
    if (!o) return;
    toggle(o);
    for (const r of box.children) r.classList.toggle("is-cur", r === o);
  });
  box.addEventListener("keydown", (e) => {
    const rows = [...box.children];
    let i = rows.findIndex((r) => r.classList.contains("is-cur"));
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      i = Math.max(0, Math.min(rows.length - 1, i + (e.key === "ArrowDown" ? 1 : -1)));
      rows.forEach((r, j) => r.classList.toggle("is-cur", j === i));
      if (rows[i]) rows[i].scrollIntoView({ block: "nearest" });
    } else if (e.key === " " && rows[i]) { e.preventDefault(); toggle(rows[i]); }
  });
}

function initTabs() {
  const tabs = [...document.querySelectorAll("#subtabs [data-tab]")];
  let cur = "strain";
  try { cur = localStorage.getItem("kensa.subtab") || cur; } catch (e) { /* 保存できなくても動く */ }
  const show = (name) => {
    for (const t of tabs) t.setAttribute("aria-selected", String(t.dataset.tab === name));
    for (const p of document.querySelectorAll(".subpane")) p.hidden = p.dataset.pane !== name;
    try { localStorage.setItem("kensa.subtab", name); } catch (e) { /* 同上 */ }
  };
  for (const t of tabs) t.addEventListener("click", () => { show(t.dataset.tab); if (t.dataset.tab === "bad") fitDefectPreview(); refocus(); });
  if (cur === "grid") cur = "strain";
  show(tabs.some((t) => t.dataset.tab === cur) ? cur : "strain");
  // 画面が低い(1366x768 など)ときは、測定表と 歪・外観… を1つのタブにまとめてスクロールをなくす。
  // 「測定表」タブは そのときだけ出る。測定器から値が来たら測定表へ戻す
  const mq = window.matchMedia("(max-height: 900px)");
  const kensa = $("kensa");
  // 2026-11-12: 表示の選び方(この PC で覚える)。all = 測定表も 歪・外観… と同じタブの 1 つにする
  const tabbed = () => mq.matches || layoutMode() === "all";
  const apply = () => {
    kensa.classList.toggle("kensa--compact", mq.matches);
    kensa.classList.toggle("kensa--tabbed", tabbed());
    kensa.classList.toggle("kensa--alltab", layoutMode() === "all");
    if (!tabbed()) {
      kensa.dataset.ctab = "";
      if (tabs.some((t) => t.dataset.tab === "grid" && t.getAttribute("aria-selected") === "true")) show("strain");
      return;
    }
    const sel = tabs.find((t) => t.getAttribute("aria-selected") === "true");
    kensa.dataset.ctab = kensa.dataset.ctab || "grid";
    if (kensa.dataset.ctab === "grid") for (const t of tabs) t.setAttribute("aria-selected", String(t.dataset.tab === "grid"));
    else if (sel && sel.dataset.tab === "grid") kensa.dataset.ctab = "grid";
  };
  for (const t of tabs) t.addEventListener("click", () => {
    if (!tabbed()) return;
    kensa.dataset.ctab = t.dataset.tab === "grid" ? "grid" : "sub";
    if (t.dataset.tab === "grid") for (const x of tabs) x.setAttribute("aria-selected", String(x === t));
  });
  showGridTab = () => {
    if (!tabbed() || kensa.dataset.ctab === "grid") return;
    kensa.dataset.ctab = "grid";
    for (const x of tabs) x.setAttribute("aria-selected", String(x.dataset.tab === "grid"));
  };
  mq.addEventListener("change", apply);
  const btn = $("layoutToggle");
  const label = () => {
    btn.setAttribute("aria-pressed", String(layoutMode() === "all"));
    btn.textContent = layoutMode() === "all" ? "表示: すべてタブ" : "表示: 測定+タブ";
  };
  btn.addEventListener("click", () => {
    try { localStorage.setItem(LAYOUT_KEY, layoutMode() === "all" ? "split" : "all"); } catch { /* 覚えられなくても動く */ }
    layoutOverride = layoutOverride === "all" ? "split" : "all";
    if (!tabbed()) kensa.dataset.ctab = "";
    else kensa.dataset.ctab = "grid";
    apply();
    label();
    refocus();
  });
  apply();
  label();
}
const LAYOUT_KEY = "kensa.layout";
let layoutOverride = null;              // localStorage が使えないときも、この画面の間は切り替わる
function layoutMode() {
  if (layoutOverride) return layoutOverride;
  try { layoutOverride = localStorage.getItem(LAYOUT_KEY) === "all" ? "all" : "split"; } catch { layoutOverride = "split"; }
  return layoutOverride;
}
let showGridTab = () => {};

// 右の列のタブ(操作・公差・チェック・Min/Max)
function initRailTabs() {
  const tabs = [...document.querySelectorAll("#rtabBar [data-rtab]")];
  const show = (name) => {
    for (const t of tabs) t.setAttribute("aria-selected", String(t.dataset.rtab === name));
    for (const p of document.querySelectorAll(".rpane")) p.hidden = p.dataset.rpane !== name;
    try { localStorage.setItem("kensa.railtab", name); } catch (e) { /* 保存できなくても動く */ }
  };
  for (const t of tabs) t.addEventListener("click", () => { show(t.dataset.rtab); refocus(); });
  let cur = "ops";
  try { cur = localStorage.getItem("kensa.railtab") || cur; } catch (e) { /* 同上 */ }
  show(tabs.some((t) => t.dataset.rtab === cur) ? cur : "ops");
}

/* 測定器はキーボードと同じで「カーソルのある欄」に打つ。入力欄・選択欄の外を押したあとでも
   受け取れるように、文字キーが来たら受け取る欄へ戻す。受け取れるかどうかは色で出す */
function initReceiver() {
  const box = $("measIn");
  const show = () => {
    if ($("recvBox").dataset.direct === "1") return;     // 受信機から直接のときはカーソルに関係ない
    const a = document.activeElement;
    const other = a && a !== box && a.matches && a.matches("input, select, textarea, [contenteditable]");
    $("recvBox").dataset.on = other ? "" : "1";
    text($("recvState"), other ? "ほかの欄に入力中 — ここを押すと測定を受けます" : "受信待ち(測定器のボタンで入ります)");
  };
  initReceiverShow = show;
  box.addEventListener("focus", show);
  box.addEventListener("blur", () => setTimeout(() => {
    // 2026-11-11: 入力欄でない所(表・ボタンなど)へ焦点が移ったら、受け取る欄へ戻す(測定中だけ。窓が開いていれば戻さない)
    const a = document.activeElement;
    const other = a && a !== document.body && a.matches("input, select, textarea, [contenteditable]");
    if (phase === "measure" && !other && !document.querySelector("dialog[open]")) box.focus({ preventScroll: true });
    show();
  }, 0));
  document.addEventListener("focusin", () => setTimeout(show, 0));
  document.addEventListener("keydown", (e) => {
    const t = e.target;
    const typing = t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT" || t.isContentEditable);
    if (typing || e.ctrlKey || e.altKey || e.metaKey) return;
    if (e.key.length === 1 && !document.querySelector("dialog[open]")) box.focus();   // この文字から受け取る欄に入る
  }, true);
  show();
}

/* 受信機から直接受け取る(保守で「検査画面でこの受信機から受け取る」を入れたとき)。
   値はサーバが受け取って検査に渡すので、ここは届いた結果を出すだけ(recv.js の startReceiver・fromServer)。
   2026-11-09: 受信機の状態・最後に届いた値と時刻・入らなかった理由を出す(「検査できない理由を表示して」) */
function recvShow(st) {
  $("recvBox").dataset.direct = st.on ? "1" : "";
  if (st.on) {
    $("recvBox").dataset.on = "1";
    text($("recvState"), `受信機から直接受信中(${st.port})— カーソルに関係なく入ります`);
    text($("recvDev"), `受信機 ${st.port}: 受信中`);
  } else {
    text($("recvDev"), `受信機: 止まっている — ${st.why || st.error || "理由不明"}(キー入力で受け取ります)`);
    initReceiverShow();
  }
  const last = st.received || {};                 // F6(2026-11-10): サーバは recv / recv/start とも received
  text($("recvLastText"), last.text ? `最後に届いた値 ${last.text}(${last.at})` + (last.screen && last.screen !== "kensa" ? ` → ${last.screen} 画面` : "") : "まだ値は届いていません");
  const rej = last.text && last.screen === "kensa" && last.why ? `入らなかった理由: ${last.why}` : "";
  $("recvWhy").hidden = !rej;
  text($("recvWhy"), rej);
}
let initReceiverShow = () => {};
function initDirect() {
  startReceiver("kensa", fromServer(op, (r) => { remember(`受信機 ${r.text}`); showGridTab(); }), recvShow);
}

function init() {
  buildExtras();
  initReceiver();
  initDirect();
  $("kAddClose").addEventListener("click", () => { $("kAdd").close(); refocus(); });
  $("kAdd").addEventListener("close", clearAddInputs);   // Esc・閉じる・受け付けのどれでも、打ちかけを残さない
  $("whoToggle").addEventListener("click", () => {
    try { localStorage.setItem(WHO_KEY, showWho() ? "0" : "1"); } catch { /* 覚えられなくても動く */ }
    $("whoToggle").setAttribute("aria-pressed", String(showWho()));
    if (V) render(V);
  });
  $("whoToggle").setAttribute("aria-pressed", String(showWho()));
  initTabs();
  initRailTabs();
  const closeAdd = () => { if ($("kAdd").open) $("kAdd").close(); };
  onEnter("cardIn", async (card) => {
    closeAdd();
    const res = await op("worker", { card });
    if (toCalibrationAfterWorker(res)) return;
    refocus();
  }, 13);
  onEnter("lotIn", async (lot) => { closeAdd(); await op("lot", { lot }); refocus(); }, 7);
  $("lotReloadBtn").addEventListener("click", async () => { await reloadLot(); refocus(); });
  $("lotWatchGo").addEventListener("click", async () => { await reloadLot(); refocus(); });
  setInterval(watchLot, WATCH_MS);
  $("lotSrcBtn").addEventListener("click", () => {
    const on = !(V && V.lot_source && V.lot_source.on);
    op("lot-source", { on, lot: $("lotIn").value });
  });
  onEnter("measIn", async (t) => { remember(t); showGridTab(); await op("receive", { text: t }); refocus(); });
  $("yamaSel").addEventListener("change", () => op("yama", { yama: $("yamaSel").value }));
  $("atBtn").addEventListener("click", () => op("at"));
  $("inspeBtn").addEventListener("click", () => op("inspe"));
  $("byousyaChk").addEventListener("change", () => op("value", { name: "Byousya", value: $("byousyaChk").checked ? "True" : "False" }));
  $("mmBtn").addEventListener("click", () => op("minmax"));
  $("printBtn").addEventListener("click", async () => {
    // 2 度押しで保存用・測定DB に二重に登録しないよう、終わるまで押せなくする(VBA の PrintB.Enabled = False)
    const b = $("printBtn");
    if (b.disabled) return;
    b.disabled = true;
    try {
      const r = await op("print-check");
      if (r && !r.stop && r.view && r.view.print_ready) await op("print-commit");
    } finally { b.disabled = false; }
  });
  $("printResetBtn").addEventListener("click", () => op("print-reset"));
  $("newBtn").addEventListener("click", () => op("new"));
  $("holdBtn").addEventListener("click", async () => { await op("hold"); refocus(); });
  $("abnoBtn").addEventListener("click", async () => { await op("value", { name: "Abno", value: "" }); refocus(); });
  $("izyouBtn").addEventListener("click", () => op("izyou"));
  $("resizeBtn").addEventListener("click", () => (V && V.buttons && V.buttons.resize ? openResize() : op("resize")));
  $("resizeOpen").addEventListener("click", openResize);
  $("kResizeClose").addEventListener("click", () => { $("kResize").close(); refocus(); });
  $("kVcClose").addEventListener("click", () => { $("kVc").close(); refocus(); });
  $("kVc").addEventListener("click", (e) => { if (e.target === $("kVc")) { $("kVc").close(); refocus(); } });
  $("aorbBtn").addEventListener("click", async () => { await op("value", { name: "AorB", value: "" }); refocus(); });
  // クリックはダブルクリックか分かるまで少し待つ。すぐ送ると描き直しで表が入れ替わり、
  // 2 回目のクリックが別の要素に当たってダブルクリック(消去)が効かないことがあった
  let clickTimer = null;
  $("mgrid").addEventListener("click", async (e) => {
    // 2026-11-11: 許容・保留の印(判定の行・見出し)を押すと外す(NG の値は消さない。サーバが確かめる)
    const flag = e.target.closest(".ngflag--allow[data-pos], .ngflag--hold[data-pos]");
    if (flag) { clearTimeout(clickTimer); await op("ng-release", { pos: flag.dataset.pos }); refocus(); return; }
    const td = e.target.closest("td[data-name]");
    if (!td) return;
    const name = td.dataset.name;
    clearTimeout(clickTimer);
    clickTimer = setTimeout(async () => { await op("click", { name }); refocus(); }, 260);
  });
  $("mgrid").addEventListener("dblclick", async (e) => {
    clearTimeout(clickTimer);
    const ng = e.target.closest("td.ng[data-pos]");
    if (ng && !e.target.closest(".ngflag--allow, .ngflag--hold")) { await op("ng-delete", { pos: ng.dataset.pos }); return; }
    const td = e.target.closest("td[data-name]");
    if (!td || !td.textContent) return;
    await op("dblclick", { name: td.dataset.name });
    refocus();
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
    // 2026-11-11: 枚数は Enter を押さずに「反映」を押されることが多い。画面に打った枚数を先に送る
    // (送らないとサーバは空のまま →「枚数入力がありません」)。VBA では ?Bad_Change で打った時点で入っていた
    const sent = Object.fromEntries(((V && V.bad) || []).map((c) => [c.pos, c.value ?? ""]));
    for (const inp of $("badGrid").querySelectorAll("input[data-name]")) {
      const pos = inp.dataset.name[0];
      if (inp.value.trim() !== "" && inp.value !== sent[pos]) await op("value", { name: inp.dataset.name, value: inp.value });
    }
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
