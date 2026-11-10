/*
  Aインプット画面(VBA の UFAinput)。
  判断はすべてサーバ(ufkensa/kensa/ainput_session.py → ainput.py)。ここは受け取った view を描き、操作を送るだけ。

  操作の返り値(Result): { view, messages, question, stop } ── 検査画面と同じ
    messages … 知らせ(VBA の MsgBox の原文)。トーストで出す。音は鳴らす
    question … 聞くこと(「ｱｲｶｽ転送を行いますか？」など)。確認窓で答えを聞き、answers に入れて呼び直す
    stop     … 止めた理由。状態は保ったまま(D10)
*/

import { api } from "../api.js";
import { toast, toastError } from "../toast.js";
import { playMessage } from "../sound.js";
import { confirmBox } from "../askbox.js";
import { startReceiver } from "../recv.js";

const $ = (id) => document.getElementById(id);
const LEVEL_KIND = { info: "info", warn: "warn", error: "error" };
let V = null;

async function op(name, body = {}, answers = {}) {
  let res;
  try {
    res = await api.post(`/api/ainput/${name}`, { ...body, answers });
  } catch (e) {
    toastError(e);
    return null;
  }
  for (const m of res.messages || []) {
    playMessage(m);                                               // 音(D29: 出来事の鍵。鳴らせなければ文言)
    if (m.text) toast(m.text, LEVEL_KIND[m.level] || "info");
  }
  if (res.view) render(res.view);
  if (res.question) {
    const q = res.question;
    const yes = await confirmBox(q.text, { ok: q.buttons === "okcancel" ? "OK" : "はい",
                                           cancel: q.buttons === "okcancel" ? "キャンセル" : "いいえ" });
    return op(name, body, { ...answers, [q.key]: !!yes });
  }
  if (res.stop) toast(res.stop.text, LEVEL_KIND[res.stop.level] || "error");
  return res;
}

function text(el, t) { el.textContent = t ?? ""; }

function kv(dl, pairs) {
  dl.replaceChildren(...pairs.filter((p) => p).flatMap(([k, v]) => {
    const dt = document.createElement("dt"); dt.textContent = k;
    const dd = document.createElement("dd"); dd.textContent = v ?? "";
    return [dt, dd];
  }));
}

function cell(t, cls) {
  const td = document.createElement("td");
  td.textContent = t ?? "";
  if (cls) td.className = cls;
  return td;
}

function off(el, hidden) { el.dataset.off = hidden ? "1" : ""; }

/** 入力中の欄は書き換えない(打っている途中の値を消さない)。 */
function setField(el, value) {
  if (document.activeElement === el && el.tagName !== "SELECT") return;
  if (el.tagName === "SELECT" && value !== "" && ![...el.options].some((o) => o.value === value)) {
    el.append(new Option(value, value));
  }
  el.value = value ?? "";
}

// ------------------------------------------------------------------
function renderOpen(v) {
  $("openBox").hidden = v.ready;
  $("body").hidden = !v.ready;
  const seg = $("kenSeg");
  seg.replaceChildren(...v.ken_methods.map((m) => {
    const b = document.createElement("button");
    b.type = "button"; b.className = "ain__method";
    const name = document.createElement("b"); name.textContent = m;
    const note = document.createElement("small"); note.textContent = KEN_NOTE[m] || "";
    b.append(name, note);
    if (v.ken_m === m) {
      const last = document.createElement("span"); last.className = "ain__last"; last.textContent = "前回";
      b.append(last);
    }
    b.setAttribute("aria-pressed", String(v.ken_m === m));
    b.addEventListener("click", () => op("open", { ken_m: m, source: V ? V.source : FROM }));
    return b;
  }));
  const k = v.kensa;
  kv($("openKv"), [["検査中のロット", k.lot || "(なし)"], ["ライン", k.line || "-"], ["Min/Max", k.minmax ? "済み" : "まだ"]]);
  const mm = $("openKv").querySelectorAll("dd")[2];
  if (mm) mm.dataset.state = k.minmax ? "ok" : "ng";
}

/** 検査法ボタンの説明(画面の違いだけ。決まりは VBA の UF1 Inspection1〜4 のまま) */
const KEN_NOTE = {
  L1: "検査法 L1 で開く(風袋・総重量の欄なし)",
  L2: "検査法 L2 で開く(風袋・総重量の欄なし)",
  M6: "検査法 M6 で開く(風袋・総重量を入れる)",
  M2: "検査法 M2 で開く(風袋・総重量を入れる)",
};

function renderLeft(v) {
  const l = v.lot;
  text($("kenLv"), `検査法 ${l.ken_law}`);
  kv($("lotKv"), [["ライン", l.line], ["Lot", l.lot], ["検査No", l.ken_no], ["材質/調質", [l.zai, l.tyo].filter(Boolean).join(" ")],
                  ["寸法", [l.atu, l.haba, l.take].filter(Boolean).join(" × ")], ["用途", l.youto],
                  ["ｵｰﾀﾞｰNo", l.order_no], l.adjust_visible ? [l.adjust_label, l.adjust_no] : null,
                  ["厚 Min〜Max", l.amin && `${l.amin} 〜 ${l.amax}`], ["巾 Min〜Max", l.hmin && `${l.hmin} 〜 ${l.hmax}`],
                  ["丈 Min〜Max", l.tmin && `${l.tmin} 〜 ${l.tmax}`]]);
  const warns = [];
  if (l.db_flag) warns.push("Lot情報履歴に履歴が無いため A_input転送は使用できません。");
  if (l.other_lot) warns.push(`検査画面のロット(${v.kensa.lot})と違います。「開き直す」で写し直せます。`);
  $("lotWarn").hidden = warns.length === 0;
  text($("lotWarn"), warns.join("\n"));

  const sel = $("f_hiki_no");
  const opts = ["", ...v.hiki_list];
  if (sel.options.length !== opts.length || opts.some((o, i) => sel.options[i].value !== o)) {
    sel.replaceChildren(...opts.map((o) => new Option(o || "- 選ぶ -", o)));
  }
  for (const [name, value] of Object.entries(v.fields)) {
    const el = $(`f_${name}`);
    if (el) setField(el, value);
  }
  const s = v.show;
  text($("maiLabel"), v.mai_label);
  document.querySelectorAll(".row2").forEach((el) => off(el, !s.mai2));
  text($("maiSpin"), s.mai2 ? "2段目を消す" : "2段目を足す");
  $("maiSpin").dataset.dir = s.mai2 ? "up" : "down";
  document.querySelectorAll(".matl").forEach((el) => off(el, !s.material));
  off($("f_material2"), !s.material2); off($("f_material3"), !s.material3);
  $("exCard").hidden = !s.ex;
  $("l_ex_case_no").hidden = $("f_ex_case_no").hidden = s.ex_case_no === false;    // EX のストアはケースNo 無し
  off($("f_ex_material2"), !s.ex_material2); off($("f_ex_material3"), !s.ex_material3);
}

function renderCenter(v) {
  const m = v.mount;
  setField($("mountIn"), m.count);
  $("deviceLv").textContent = m.device || "未表示";
  $("deviceLv").className = "lv " + (m.device ? "lv--info" : "");
  text($("perMount"), m.per_mount && `理論値(1山) = 厚×枚数1×包数1÷山数 = ${m.per_mount} mm`);
  off($("measBox"), !m.visible);
  text($("hint"), m.hint);
  const r = m.received;
  const rec = $("received");
  rec.className = r.reason ? "ain__received-ng" : "muted";
  text(rec, r.text ? (r.reason ? `受け付けませんでした: ${r.text} … ${r.reason}` : `${r.index} 山目に入れました: ${r.text}`) : "");

  const body = $("heights").tBodies[0];
  body.replaceChildren(...(m.visible ? m.heights : []).map((h) => {
    const tr = document.createElement("tr");
    tr.dataset.next = h.next ? "1" : "";
    const x = document.createElement("button");
    x.type = "button"; x.className = "x"; x.textContent = "×"; x.title = `${h.n} 山目を消す`;
    x.hidden = !h.value;
    x.addEventListener("click", () => op("clear-height", { n: h.n }));
    const last = document.createElement("td"); last.append(x);
    tr.append(cell(`${h.n} 山目`), cell(m.per_mount), cell(h.value, "val"), last);
    return tr;
  }));
  const foot = $("heights").tFoot;
  foot.replaceChildren();
  if (m.visible && m.heights.length) {
    const tr = document.createElement("tr");
    tr.append(cell(`合計 ${m.filled}/${m.heights.length}`), cell(""), cell(m.total, "val"), cell(""));
    foot.append(tr);
  }

  const c = v.check;
  $("checkCard").hidden = !c;
  if (c) {
    $("checkLv").textContent = c.judged ? (c.ok ? "OK" : "NG") : "判定なし";
    $("checkLv").className = "lv " + (c.judged ? (c.ok ? "lv--ok" : "lv--ng") : "lv--warn");
    const msg = $("checkMsg");
    msg.dataset.ng = c.ok ? "" : "1";
    text(msg, c.ok ? c.label : c.msg);
    kv($("checkKv"), c.judged ? [
      ["実測(合計)", `${c.hei} mm`],
      ["理論値", `${c.gm} 〜 ${c.gp} mm(Min厚〜Max厚 × 枚数 + 合紙・VC・外装紙)`],
      ["合紙 / VC / 外装紙", `${c.ais} / ${c.vcs} / ${c.gai}`],
      c.ok ? ["許容", `上 +${c.allow_over}(最大厚×3/4)/ 下 −${c.allow_short}(最小厚×1/3)`]
           : ["外れ", `${c.diff} mm(許容 ${c.allow})`],
      ["1山の目安", `${c.lo} 〜 ${c.hi} mm(理論値 ÷ 山数。判定は合計だけ)`],
    ] : []);
    const tb = $("checkRows").tBodies[0];
    $("checkRows").hidden = !c.judged || c.ok;
    tb.replaceChildren(...(c.rows || []).map((r) => {
      const tr = document.createElement("tr");
      tr.dataset.out = r.out ? "1" : "";
      tr.dataset.worst = r.worst ? "1" : "";
      tr.append(cell(`${r.n} 山目`), cell(r.value, "val"), cell(r.diff ? `${r.diff} mm` : "目安内"));
      return tr;
    }));
  }

  const x = v.extras;
  for (const [id, key] of [["x_ai", "ai"], ["x_vab", "vab"], ["x_ga", "ga"]]) {
    $(id).checked = !!x[key];
    $(id).disabled = x.locked;
  }
  text($("aisi"), x.aisi); text($("vcab"), x.vcab); text($("gaisousi"), x.gaisousi);
  $("gaityuuBox").hidden = !v.show.gaityuu;
  $("unlockBtn").disabled = !x.locked;

  text($("pressText"), v.press.text);
  $("pressBtn").disabled = v.busy;
}

const STAGE_LV = { sent: ["送信済み", "lv--ok"], stopped: ["停止", "lv--ng"], height_ng: ["高さNG", "lv--ng"],
                   input_error: ["入力不足", "lv--warn"], cancel: ["取りやめ", "lv--warn"], question: ["確認待ち", "lv--info"],
                   first: ["入力欄を表示", "lv--info"], reset: ["Reset", "lv--info"] };

function renderRight(v) {
  const s = v.send;
  const [lab, cls] = STAGE_LV[s.stage] || ["まだ", ""];
  $("sendLv").textContent = lab;
  $("sendLv").className = "lv " + cls;
  $("steps").replaceChildren(...s.steps.map((st) => {
    const li = document.createElement("li");
    li.dataset.state = st.state;
    li.textContent = st.label;
    return li;
  }));
  const res = $("sendLabel");
  res.dataset.state = s.stage === "sent" ? (s.result_missing ? "" : "ok") : (["stopped", "height_ng", "input_error"].includes(s.stage) ? "ng" : "");
  text(res, s.label);
  kv($("sendKv"), [
    s.at ? ["時刻", s.at] : null,
    s.csv_path ? ["梱包課CSV", s.csv_path] : null,
    s.hostsend_path ? ["hostsend", s.hostsend_path] : null,
    ["送信アプリ", s.sender],
    s.ng_flag ? ["Ngflag", `${s.ng_flag}(開き直すと解けます)`] : null,
  ]);
  kv($("codes"), s.codes.map((c) => [c.code, c.text]));
  $("sends").replaceChildren(...s.history.map((h) => {
    const li = document.createElement("li");
    li.textContent = `${h.at} ${h.stage === "sent" ? "送信" : "停止"} 引当${h.hiki_no} ${h.note}`;
    return li;
  }));
  if (!s.history.length) { const li = document.createElement("li"); li.textContent = "まだ送っていません"; $("sends").append(li); }
  text($("current"), s.current);
}

function render(v) {
  V = v;
  renderOpen(v);
  if (!v.ready) return;
  renderLeft(v);
  renderCenter(v);
  renderRight(v);
}

// ------------------------------------------------------------------
const QS = new URLSearchParams(location.search);
const FROM = QS.get("from") || "";          // 呼び元(Cask / Mitsubishi)。無ければ通常検査
const METHOD = QS.get("method") || "";

function init() {
  document.querySelectorAll("[data-field]").forEach((el) => {
    const send = () => op("value", { name: el.dataset.field, value: el.type === "checkbox" ? (el.checked ? "1" : "0") : el.value });
    el.addEventListener("change", send);
    if (el.tagName === "INPUT" && el.type !== "checkbox") {
      el.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); el.blur(); } });
    }
  });
  document.querySelectorAll("[data-spin]").forEach((b) => {
    b.addEventListener("click", () => op("spin", { name: b.dataset.spin, dir: b.dataset.dir }));
  });
  $("maiSpin").addEventListener("click", () => op("spin", { name: "mai", dir: $("maiSpin").dataset.dir }));
  const mount = $("mountIn");
  const sendMount = () => { if (mount.value !== "" && V && mount.value !== V.mount.count) op("mount", { value: mount.value }); };
  mount.addEventListener("change", sendMount);
  mount.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); sendMount(); $("measIn").focus(); } });
  // 測定器の伝送: VBA は Tab で受け取る(StartTB_KeyDown)。Enter でも受ける
  $("measIn").addEventListener("keydown", async (e) => {
    if (e.key !== "Enter" && e.key !== "Tab") return;
    e.preventDefault();
    const t = $("measIn").value;
    $("measIn").value = "";
    if (t) await op("receive", { text: t });
    $("measIn").focus();
  });
  $("pressBtn").addEventListener("click", async () => {
    await op("press");
    if (V && V.ready && V.mount.visible && V.mount.next > 0) $("measIn").focus();
  });
  $("unlockBtn").addEventListener("click", () => op("unlock"));
  $("resetBtn").addEventListener("click", () => op("reset"));
  $("reopenBtn").addEventListener("click", () => op("open", { ken_m: V ? V.ken_m : "", source: V ? V.source : FROM }));
  api.get("/api/ainput/state").then((res) => {
    render(res.view);
    const v = res.view;
    // キャスク・全ヵ所の「Aインプットへ」から来た: 呼び元のシートのロット・Min/Max で開く(Input:247-252)
    if (FROM) return op("open", { ken_m: METHOD || v.ken_m, source: FROM });
    // 開いていない・検査画面のロットが変わった → 前に選んだ検査法で開き直す(VBA は毎回 Ainput_Startup)
    if (v.ken_m && (!v.ready || (v.kensa.lot && v.lot.lot !== v.kensa.lot && !v.lot.db_flag))) return op("open", { ken_m: v.ken_m });
    return null;
  }).catch(toastError);
}

init();

// 受信機から直接受け取る(保守で「検査画面で受け取る」にした受信機)。この画面が開いている間だけ、この画面に入る
startReceiver("ainput", (text) => op("receive", { text }));
