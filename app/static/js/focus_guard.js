/* 2026-11-24(現場の声): 画面の文字をドラッグで選んでコピーできるように。
   測定器の値を受ける欄へ焦点を戻す処理が、選んだ文字を毎回消していた。
   文字を選んでいる間・マウスを押している間は、焦点を戻さない(測定器の文字キーが来たら戻る)。 */
let pressed = false;
document.addEventListener("pointerdown", () => { pressed = true; }, { capture: true, passive: true });
document.addEventListener("pointerup", () => { pressed = false; }, { capture: true, passive: true });
document.addEventListener("pointercancel", () => { pressed = false; }, { capture: true, passive: true });

/** いま焦点を受け取る欄へ戻してよいか(文字を選んでいない・押していない)。 */
export function mayRefocus() {
  if (pressed) return false;
  const sel = window.getSelection ? window.getSelection() : null;
  return !(sel && !sel.isCollapsed && String(sel).trim() !== "");
}
