// 操作できる画面は1つだけ(2026-11-11)。tests/e2e/test_e2e.py から node で動かす(最後に)。
// 2つ目のタブ → 覆い「この画面はほかのウィンドウで開いています」。「こちらで使う」で取り返す → 1つ目に覆い。
// 使っていない画面からの書く操作はサーバが断る(inactive_screen)。2回目の起動(/api/screen/raise)で知らせが出る。
const { BASE, setup, close, shot, check, finish } = require("./common");

(async () => {
  const { browser, ctx, page, S } = await setup();
  const W = (ms) => page.waitForTimeout(ms);
  const locked = (p) => p.evaluate(() => { const el = document.getElementById("screenLock"); return !!el && !el.hidden; });
  // その画面の番号を付けて書く操作を送る
  const write = (p) => p.evaluate(async () => {
    const r = await fetch("/api/e2e-no-such-write", { method: "POST", headers: {
      "Content-Type": "application/json", "X-Tool-Token": window.APP.token, "X-Screen-Id": window.APP.screenId }, body: "{}" });
    let j = null; try { j = await r.json(); } catch { /* 空 */ }
    return { status: r.status, code: j && j.error && j.error.code };
  });
  const waitLocked = async (p, want, ms = 8000) => {
    const end = Date.now() + ms;
    while (Date.now() < end) { if ((await locked(p)) === want) return true; await p.waitForTimeout(250); }
    return false;
  };
  let page2 = null;
  try {
    S.step = "single";
    // 前の試験の画面が閉じきるのを待つ(閉じたことが届かなかった画面は、心拍が途絶えて10秒で外れる)
    await page.goto(BASE + "/");
    check(await waitLocked(page, false, 15000), S.step, "1つ目の画面は操作できる");

    page2 = await ctx.newPage();
    page2.on("pageerror", (e) => S.errors.push(`[${S.step}] page2 ${e.message}`));
    await page2.goto(BASE + "/");
    check(await waitLocked(page2, true), S.step, "2つ目の画面に覆いが出る");
    const text = await page2.locator("#screenLock").textContent();
    check(/この画面はほかのウィンドウで開いています/.test(text) && /こちらで使う/.test(text) && /閉じる/.test(text),
      S.step, "覆いの文と 2 つのボタン");
    let r = await write(page2);
    check(r.status === 409 && r.code === "inactive_screen", S.step, "使っていない画面の書く操作は断る " + JSON.stringify(r));
    r = await write(page);
    check(r.status === 404, S.step, "使っている画面の書く操作は通る(経路が無いので 404) " + JSON.stringify(r));
    await shot(page2, "single_screen_overlay");

    // 取り返す
    await page2.click('#screenLock [data-screen="take"]');
    await page2.waitForLoadState("load"); await page2.waitForTimeout(1200);
    check(await waitLocked(page2, false), S.step, "「こちらで使う」で 2つ目が操作できる");
    check(await waitLocked(page, true), S.step, "1つ目に覆いが出る");
    r = await write(page);
    check(r.status === 409 && r.code === "inactive_screen", S.step, "取られた画面の書く操作は断る " + JSON.stringify(r));
    r = await write(page2);
    check(r.status === 404, S.step, "取り返した画面の書く操作は通る " + JSON.stringify(r));

    // 読み込み直しても使っている画面のまま
    await page2.reload(); await page2.waitForTimeout(1500);
    check(await waitLocked(page2, false), S.step, "読み込み直しても使っている画面のまま");
    check(await locked(page), S.step, "1つ目は覆いのまま");

    // 2回目の起動: 開いている画面に「前に出て」
    const raised = await page2.evaluate(async () => (await (await fetch("/api/screen/raise", { method: "POST" })).json()));
    check(raised.raised === true, S.step, "開いている画面があれば raised(ブラウザを開かない)");
    let toasted = false;
    for (let i = 0; i < 24 && !toasted; i++) {
      await page2.waitForTimeout(250);
      toasted = /すでに開いています/.test(await page2.evaluate(() => document.getElementById("toasts").textContent || ""));
    }
    check(toasted, S.step, "前に出た画面に「すでに開いています(もう一度起動されました)」");
    await shot(page2, "single_screen_raised");
  } catch (e) {
    check(false, S.step, "例外 " + (e && e.stack || e));
  }
  finish(S);
  await close(browser);
})();
