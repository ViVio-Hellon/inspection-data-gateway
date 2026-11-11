/*
  judgebox.js — 検査画面の大きな知らせ(2026-11-09)

  1. 異常値の窓(VBA の UF異常値発見 フォーム)
     トーストでは小さく見落とすので、画面いっぱいの赤い窓で出し、「確認」を押すまで閉じない。
     中身はサーバの知らせ(title「異常」)の文: 1 行目が値、そのあとが閾値(公差の 2.0 倍など)。

  2. 判定の画像(VBA の UFOK / UFNG / UF対角 フォームの絵)
     合格・不合格・対角線の出来事(音と同じ鍵 cue)で、設定 → 使用パス「判定の画像」の画像を大きく出す。
     OK・NG は 3 秒、対角線は 2 秒(VBA と同じ。保守 → 判定の画像 で時間・大きさを PC ごとに変えられる)。押すか Esc でもすぐ消える。
     フォルダが空・ファイルが無いときは何も出さない(エラーにしない)。
*/
import { api, tokenUrl } from "./api.js";

let styled = false;
function style() {
  if (styled) return;
  styled = true;
  const s = document.createElement("style");
  s.textContent = `
    .alarmbox { border: 4px solid #c62828; border-radius: var(--radius, 10px); padding: 0;
                width: min(760px, 94vw); background: #fff5f5; color: #3b0d0d;
                box-shadow: 0 25px 80px rgb(120 0 0 / 45%); }
    .alarmbox::backdrop { background: rgb(120 0 0 / 55%); }
    .alarmbox__head { background: #c62828; color: #fff; font-size: 34px; font-weight: 800;
                      letter-spacing: .08em; padding: 14px 24px; }
    .alarmbox__value { font-size: 52px; font-weight: 800; color: #b71c1c; padding: 18px 24px 4px;
                       font-variant-numeric: tabular-nums; word-break: break-all; }
    .alarmbox__lines { margin: 0; padding: 4px 24px 8px 44px; font-size: 20px; line-height: 1.6; }
    .alarmbox__foot { display: flex; justify-content: flex-end; padding: 12px 24px 20px; }
    .alarmbox__ok { font-size: 24px; font-weight: 700; padding: 12px 40px; border-radius: 10px;
                    border: 0; background: #c62828; color: #fff; cursor: pointer; }
    .alarmbox__ok:focus-visible { outline: 4px solid #ffb300; outline-offset: 3px; }
    .packbox { border: 4px solid #ef6c00; border-radius: var(--radius, 10px); padding: 0;
               width: min(900px, 94vw); max-height: 92vh; background: #fffaf2; color: #2b1600;
               box-shadow: 0 25px 80px rgb(120 60 0 / 45%); }
    .packbox[open] { display: flex; flex-direction: column; }
    .packbox::backdrop { background: rgb(80 40 0 / 55%); }
    .packbox__head { background: #ef6c00; color: #fff; font-size: 30px; font-weight: 800; letter-spacing: .06em; padding: 12px 24px; }
    .packbox__body { overflow: auto; padding: 8px 24px; }
    .packbox__sec { margin: 12px 0; }
    .packbox__sec h3 { margin: 0 0 6px; font-size: 22px; color: #e65100; border-bottom: 2px solid #ffcc80; }
    .packbox__sec ul { margin: 0; padding-left: 28px; font-size: 24px; font-weight: 700; line-height: 1.5; }
    .packbox__foot { display: flex; justify-content: flex-end; padding: 12px 24px 20px; }
    .packbox__ok { font-size: 24px; font-weight: 700; padding: 12px 48px; border-radius: 10px; border: 0;
                   background: #ef6c00; color: #fff; cursor: pointer; }
    .packbox__ok:focus-visible { outline: 4px solid #1565c0; outline-offset: 3px; }
    .judgeimg { position: fixed; inset: 0; z-index: 9000; display: flex; align-items: center; justify-content: center;
                background: rgb(0 0 0 / 35%); cursor: pointer; }
    .judgeimg[hidden] { display: none; }
    .judgeimg[data-full] img { width: 100vw; height: 100vh; object-fit: contain; border-radius: 0; }
    .judgeimg img { max-width: var(--judge-size, 80vw); max-height: var(--judge-size-h, 80vh); border-radius: 12px; background: #fff;
                    box-shadow: 0 25px 80px rgb(0 0 0 / 45%); }
  `;
  document.head.appendChild(s);
}

/* ---------------- 異常値 ---------------- */
let alarm = null;
let alarmQueue = Promise.resolve();

function ensureAlarm() {
  if (alarm && document.body.contains(alarm)) return alarm;
  style();
  alarm = document.createElement("dialog");
  alarm.className = "alarmbox";
  alarm.id = "alarmBox";
  alarm.setAttribute("role", "alertdialog");
  alarm.setAttribute("aria-labelledby", "alarmHead");
  alarm.innerHTML = `
    <div class="alarmbox__head" id="alarmHead">異常値を見つけました</div>
    <div class="alarmbox__value"></div>
    <ul class="alarmbox__lines"></ul>
    <div class="alarmbox__foot"><button class="alarmbox__ok" type="button">確認しました</button></div>`;
  document.body.appendChild(alarm);
  return alarm;
}

/** 異常値の窓。「確認しました」を押すまで閉じない(Esc では閉じない)。Promise は押したときに解ける。 */
export function alarmBox(text) {
  const run = () => new Promise((resolve) => {
    const d = ensureAlarm();
    const lines = String(text || "").split(/\r?\n/).filter((x) => x !== "");
    d.querySelector(".alarmbox__value").textContent = lines[0] || "";
    d.querySelector(".alarmbox__lines").replaceChildren(...lines.slice(1).map((t) => {
      const li = document.createElement("li");
      li.textContent = t;
      return li;
    }));
    const ok = d.querySelector(".alarmbox__ok");
    const onCancel = (e) => e.preventDefault();          // Esc では閉じない(確かめてもらう)
    const done = () => {
      ok.removeEventListener("click", done);
      d.removeEventListener("cancel", onCancel);
      d.close();
      resolve(true);
    };
    ok.addEventListener("click", done);
    d.addEventListener("cancel", onCancel);
    hideImage();
    if (!d.open) d.showModal();
    ok.focus();
  });
  alarmQueue = alarmQueue.then(run);
  return alarmQueue;
}

/* ---------------- 包装の知らせ(2026-11-11) ----------------
   VBA では Judg.PackagingSpecificationNo(「注意_包装仕様」)と QP.MaterialsUpperPart(「梱包保護材」)の
   MsgBox。トーストでは見落とすので、1 つの大きな窓にまとめ、「確認」を押すまで閉じない(Esc でも閉じない)。
   sections: [{title, lines: [文…]}](文はサーバの原文のまま) */
let packQueue = Promise.resolve();
export const PACKING_TITLES = { packing_spec: "包装仕様の注意", packing_material: "梱包保護材" };
export function packingBox(sections) {
  const run = () => new Promise((resolve) => {
    style();
    let d = document.getElementById("packingBox");
    if (!d) {
      d = document.createElement("dialog");
      d.className = "packbox";
      d.id = "packingBox";
      d.setAttribute("role", "alertdialog");
      d.setAttribute("aria-labelledby", "packingHead");
      d.innerHTML = `<div class="packbox__head" id="packingHead">包装の注意</div><div class="packbox__body"></div>
        <div class="packbox__foot"><button class="packbox__ok" type="button">確認</button></div>`;
      document.body.appendChild(d);
    }
    d.querySelector(".packbox__body").replaceChildren(...sections.filter((s) => s.lines.length).map((s) => {
      const sec = document.createElement("section");
      sec.className = "packbox__sec";
      sec.dataset.kind = s.kind || "";
      const h = document.createElement("h3");
      h.textContent = s.title;
      const ul = document.createElement("ul");
      ul.append(...s.lines.map((t) => { const li = document.createElement("li"); li.textContent = t; return li; }));
      sec.append(h, ul);
      return sec;
    }));
    const ok = d.querySelector(".packbox__ok");
    const onCancel = (e) => e.preventDefault();
    const done = () => { ok.removeEventListener("click", done); d.removeEventListener("cancel", onCancel); d.close(); resolve(true); };
    ok.addEventListener("click", done);
    d.addEventListener("cancel", onCancel);
    if (!d.open) d.showModal();
    ok.focus();
  });
  packQueue = packQueue.then(run);
  return packQueue;
}

/** 包装の知らせを集めて、同じ操作の分を 1 つの窓に出す。まとめる知らせなら true。 */
let packPending = null;
export function collectPacking(m) {
  if (!m || !PACKING_TITLES[m.kind] || !m.text) return false;
  if (!packPending) {
    packPending = { packing_spec: [], packing_material: [] };
    setTimeout(() => {
      const p = packPending;
      packPending = null;
      packingBox(Object.keys(PACKING_TITLES).map((k) => ({ kind: k, title: PACKING_TITLES[k], lines: p[k] })));
    }, 0);
  }
  packPending[m.kind].push(m.text);
  return true;
}

/** サーバの知らせが異常値か(UF異常値発見: title「異常」)。 */
export function isAbnormal(m) {
  return !!m && m.title === "異常" && !!m.text;
}

/* ---------------- 判定の画像 ---------------- */
const CUE_KIND = { kensa_ok: "ok", kensa_ng: "ng", kensa_diag: "diag" };
let images = null;           // {ok: [枠], ng: [...], diag: [...]}(2026-11-09: 枠 judge_ok など)
let showMs = { ok: 3000, ng: 3000, diag: 2000 };
let sizePct = { ok: 80, ng: 80, diag: 80 };    // 2026-11-09: 画面に対する % (保守 → 判定の画像。100 = 全画面)
let layer = null, timer = 0, loadedAt = 0;

async function loadImages() {
  if (images && Date.now() - loadedAt < 30000) return images;     // 設定を変えたら 30 秒以内に効く
  try {
    const r = await api.get("/api/kensa/judge_images");
    images = r.images || {};
    showMs = { ...showMs, ...(r.show_ms || {}) };
    sizePct = { ...sizePct, ...(r.size || {}) };
  } catch (e) {
    images = images || {};
  }
  loadedAt = Date.now();
  return images;
}

function ensureLayer() {
  if (layer && document.body.contains(layer)) return layer;
  style();
  layer = document.createElement("div");
  layer.className = "judgeimg";
  layer.id = "judgeImage";
  layer.hidden = true;
  layer.innerHTML = `<img alt="">`;
  layer.addEventListener("click", hideImage);
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") hideImage(); });
  document.body.appendChild(layer);
  return layer;
}

function hideImage() {
  if (timer) clearTimeout(timer);
  timer = 0;
  if (layer) layer.hidden = true;
}

/** 合格・不合格・対角線の出来事で画像を出す。画像が無ければ何もしない。 */
export async function showJudgeImage(m) {
  const kind = m && CUE_KIND[m.cue || m.sound];
  if (!kind) return;
  const list = (await loadImages())[kind] || [];
  if (!list.length) return;
  if (document.querySelector("dialog.alarmbox[open]")) return;     // 異常値の窓を隠さない
  const name = list[Math.floor(Math.random() * list.length)];       // UF対角: 乱数で DB1〜DB5 のどれか
  display(kind, name, showMs[kind] || 3000, sizePct[kind] || 80);
}

/** 保守 → 判定の画像の「テスト表示」。検査画面と同じ窓で、時間・大きさ・枠を指定して出す(slot 空 = 乱数)。 */
export async function testJudgeImage(kind, { slot = "", ms, size } = {}) {
  const list = (await loadImages())[kind] || [];
  const name = slot || list[Math.floor(Math.random() * list.length)];
  if (!name) return "";
  display(kind, name, ms || showMs[kind] || 3000, size || sizePct[kind] || 80);
  return name;
}

function display(kind, name, ms, size) {
  const el = ensureLayer();
  const img = el.querySelector("img");
  img.alt = { ok: "OK", ng: "NG", diag: "対角線" }[kind];
  img.src = tokenUrl(`/api/kensa/judge_image/${encodeURIComponent(name)}`);   // 2026-11-09: 枠ごと
  el.style.setProperty("--judge-size", `${size}vw`);
  el.style.setProperty("--judge-size-h", `${size}vh`);
  if (size >= 100) el.dataset.full = "1"; else delete el.dataset.full;
  el.dataset.kind = kind;
  el.dataset.slot = name;
  el.hidden = false;
  if (timer) clearTimeout(timer);
  timer = setTimeout(hideImage, ms);
}
