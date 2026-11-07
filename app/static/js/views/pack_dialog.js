/*
  包装仕様書の窓(VBA の UF_PackNo。全ヵ所とキャスクの画面で使う)。
  判断はサーバ(ufkensa/kensa/packing.py)。ここは view.pack を描き、行のクリックと「閉じる」を送るだけ。

  VBA と同じくモーダル(開いている間はシートを触れない)。「確認しました」に印を付けないと閉じられない
  (付けずに閉じるとサーバが「確認後に押してください。」を返す)。

  使い方: const pack = packDialog({ imageUrl: "/api/zenkasho/pack_image", send: (op, body) => op(...) });
          pack.render(view.pack)   // null なら閉じる
*/
import { tokenUrl } from "../api.js";

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
}

export function packDialog({ imageUrl, send }) {
  const d = el("dialog", "pk");
  d.setAttribute("aria-labelledby", "pkTitle");
  d.innerHTML = `
    <header class="pk__head">
      <h3 id="pkTitle">包装仕様書</h3>
      <span class="pk__no" id="pkNo"></span>
      <span class="pk__zoom">
        <button class="btn" type="button" data-zoom="-1" aria-label="縮小">－</button>
        <button class="btn" type="button" data-zoom="1" aria-label="拡大">＋</button>
      </span>
    </header>
    <div class="pk__body">
      <div class="pk__pic"><img id="pkImg" alt="包装仕様書の画像"></div>
      <div class="pk__side">
        <table class="pk__rows" aria-label="引当">
          <thead><tr><th>引当NO</th><th>ｵｰﾀﾞｰNo</th><th>包装仕様NO</th></tr></thead>
          <tbody id="pkRows"></tbody>
        </table>
        <ul class="pk__spec" id="pkSpec"></ul>
        <label class="pk__ok"><input type="checkbox" id="pkOk"> 確認しました</label>
        <button class="btn btn--commit" type="button" id="pkClose">閉じる</button>
      </div>
    </div>`;
  document.body.append(d);
  const $ = (id) => d.querySelector(`#${id}`);
  let zoom = 1;
  let shown = "";
  d.addEventListener("cancel", (e) => e.preventDefault());          // ×(Esc)では閉じない(UF_PackNo.QueryClose)
  d.querySelectorAll("[data-zoom]").forEach((b) => b.addEventListener("click", () => {
    zoom = Math.min(2, Math.max(0.3, zoom * (b.dataset.zoom === "1" ? 1.1 : 0.9)));   // SpinButton1(±10%、0.3〜2倍)
    $("pkImg").style.width = `${Math.round(zoom * 100)}%`;
  }));
  $("pkRows").addEventListener("click", (e) => {
    const tr = e.target.closest("tr[data-i]");
    if (tr) send("pack-select", { i: Number(tr.dataset.i) });
  });
  $("pkClose").addEventListener("click", () => send("pack-close", { confirmed: $("pkOk").checked }));

  function render(p) {
    if (!p) {
      if (d.open) d.close();
      shown = "";
      return;
    }
    const key = `${p.lot}:${p.tb_load}`;
    if (key !== shown) { $("pkOk").checked = false; zoom = 1; $("pkImg").style.width = "100%"; }
    shown = key;
    $("pkTitle").textContent = `包装仕様書 ${p.caption || p.lot}`;
    $("pkNo").textContent = p.picture;
    $("pkImg").src = tokenUrl(`${imageUrl}?n=${encodeURIComponent(p.picture)}`);
    const rows = [];
    for (let i = 1; i <= 5; i += 1) {
      const r = p.rows[i - 1] || { hiki: "", ord: "", pac: "" };
      const tr = el("tr");
      tr.dataset.i = i;
      if (p.selected === i) tr.className = "sel";
      tr.append(el("td", "", r.hiki), el("td", "", r.ord), el("td", "", r.pac));
      rows.push(tr);
    }
    $("pkRows").replaceChildren(...rows);
    $("pkSpec").replaceChildren(...p.spec.filter((x) => x).map((x, i) => el("li", i === 0 ? "pk__l1" : "", x)));
    if (!d.open) d.showModal();
  }
  return { render };
}
