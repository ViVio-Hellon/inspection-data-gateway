/*
  キャスク・全ヵ所の表の入力(2026-11-26 現場の声「Excel の良いところだけ引き継ぐ」)。

  引き継ぐもの(Excel の良いところ)
  - 表のセルを押すと、そのセルが入力先になる
  - 選んだセルにそのまま打てる(セルの上に打っている文字が出る)。Enter で確定、Esc でやめる
  - 矢印キーで選ぶセルを動かす
  - Delete で選んだセルを消す / Back space(打っていないとき)は画面ごとの決まり
    (キャスク = 直前の入力の取り消し。前の「Back space」ボタンと同じ / 全ヵ所 = 選んだセルを消す)
  やめるもの(Excel の悪いところ)
  - 見える入力欄・「Back space」「この欄を消す」のボタン(キーボードでできる)
  - 日本語入力の確定の Enter で、打ちかけの文字が送られる(外観に文字が入らなかった原因)

  測定器(キーボードと同じく文字 + Enter を送る)も、この隠れた欄で受ける。
*/
const NAV = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };

/**
 * @param {object} o
 * @param {HTMLInputElement} o.input   受け取る欄(画面の外に置く)
 * @param {() => HTMLElement|null} o.cell   いま選んでいるセルの要素(打っている文字を重ねて出す)
 * @param {(text: string) => any} o.send    Enter で確定した文字
 * @param {() => any} o.backspace           打っていないときの Back space
 * @param {() => any} o.del                 Delete
 * @param {(dr: number, dc: number) => any} o.move   矢印キー
 */
export function sheetInput(o) {
  const inp = o.input;
  const tip = document.createElement("div");
  tip.className = "sheet-typing";
  tip.hidden = true;
  tip.setAttribute("aria-hidden", "true");
  document.body.append(tip);

  const showTip = () => {
    const v = inp.value;
    const c = o.cell();
    if (!v || !c) { tip.hidden = true; return; }
    const r = c.getBoundingClientRect();
    tip.textContent = v;
    tip.style.left = `${Math.round(r.left + window.scrollX)}px`;
    tip.style.top = `${Math.round(r.top + window.scrollY)}px`;
    tip.style.minWidth = `${Math.round(r.width)}px`;
    tip.style.minHeight = `${Math.round(r.height)}px`;
    tip.hidden = false;
  };
  inp.addEventListener("input", showTip);
  inp.addEventListener("compositionupdate", (e) => {
    // 変換中の文字も見せる(確定前は input.value に入らないことがある)
    const c = o.cell();
    if (!c) return;
    tip.textContent = inp.value || e.data || "";
    tip.hidden = !tip.textContent;
  });
  inp.addEventListener("blur", () => { tip.hidden = true; });
  window.addEventListener("scroll", showTip, true);

  inp.addEventListener("keydown", async (e) => {
    // 日本語入力の変換中(確定の Enter を含む)は何もしない。ここで送ると打ちかけの文字が入る
    if (e.isComposing || e.keyCode === 229) return;
    const typing = inp.value !== "";
    if (e.key === "Enter" || (e.key === "Tab" && typing)) {
      e.preventDefault();
      const t = inp.value;
      inp.value = "";
      tip.hidden = true;
      if (t !== "") await o.send(t);
      return;
    }
    if (e.key === "Escape") { inp.value = ""; tip.hidden = true; return; }
    if (typing) return;                                 // 打っている途中の Back space・矢印は文字の直し
    if (e.key === "Delete") { e.preventDefault(); await o.del(); return; }
    if (e.key === "Backspace") { e.preventDefault(); await o.backspace(); return; }
    if (NAV[e.key]) { e.preventDefault(); await o.move(...NAV[e.key]); }
  });
  return { refresh: showTip };
}
