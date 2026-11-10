// 検査画面の通しの試験(2026-11-10 リファクタリング 第1段 5)。tests/e2e/test_e2e.py から node で動かす。
// 受信機の状態 → 検査員・ロット・作業 → 公差外の確認の窓(選ぶまで閉じない)→ 受信機(疑似端末)と受信 API で測定
// → 作業範囲の選び欄(F4)→ Min/Max → 印刷・保存 → サイズ修正の窓 → 1366×768 ではみ出さない → 全ヵ所の検入高さ(F5)
const { BASE, setup, shot, askAll, api, note, check, finish } = require("./common");

const PTY = process.env.E2E_PTY || "-";
const WORKER = "9000000000001";
const LOT = "H7345B0";
const WL = { "5": "1252.3", "6": "1252.5", "7": "2502.4", "8": "2502.6", "9": "2800.1", "0": "2800.6" };

(async () => {
  const { browser, page, S } = await setup();
  const W = (ms) => page.waitForTimeout(ms);
  const feed = (text) => api(page, "POST", "/__e2e/feed", { text });
  const opAll = async (name, body = {}) => {          // 画面の op() と同じ(確認には「はい」)
    const answers = {};
    for (let i = 0; i < 8; i++) {
      const r = (await api(page, "POST", "/api/kensa/" + name, { ...body, answers })).json || {};
      if (!r.question) return r;
      answers[r.question.key] = true;
    }
    return {};
  };
  const state = async () => (await api(page, "GET", "/api/kensa/state")).json.view;
  const closeDialogs = async () => {
    for (let i = 0; i < 4; i++) {
      const d = page.locator("dialog[open]:not(#confiDlg) button").last();
      if (!(await d.count())) break;
      await d.click().catch(() => {}); await W(200);
    }
  };
  try {
    // ---- 受信機の状態 --------------------------------------------------------
    S.step = "receiver";
    await page.goto(BASE + "/kensa"); await W(1500); await askAll(page);
    check(/受信機が登録されていません/.test(await page.locator("#recvBox").innerText()), S.step, "受信機なし → 理由を出す");
    if (PTY !== "-") {
      const r = await api(page, "POST", "/api/maint/rx/add", { port: PTY });
      check(r.status === 200, S.step, "疑似端末を受信機として足す " + r.status);
      await page.reload(); await W(1500);
      check(/受信中/.test(await page.locator("#recvBox").innerText()), S.step, "受信機から直接受信中");
      await feed("DT10001+0000015.004M"); await W(1500);
      check(/DT10001\+0000015\.004M/.test(await page.locator("#recvLastText").innerText()), S.step, "最後に届いた値を出す");
      check(/入らなかった理由/.test(await page.locator("#recvWhy").innerText()), S.step, "入らなかった理由を出す");
      // F6: 開き直した直後(recv/start の返事)でも最後に届いた値が出る
      await page.reload(); await W(700);
      check(/DT10001\+0000015\.004M/.test(await page.locator("#recvLastText").innerText()), S.step, "開き直しても最後の値が出る(F6)");
    } else note(S.step, "疑似端末が無いので受信機の確認は飛ばす");

    // ---- 検査員・ロット・作業 --------------------------------------------------
    S.step = "worker";
    await page.fill("#cardIn", WORKER); await page.press("#cardIn", "Enter"); await W(800); await askAll(page);
    S.step = "lot";
    await page.fill("#lotIn", LOT); await page.press("#lotIn", "Enter"); await W(1200); await askAll(page); await W(500);
    await askAll(page); await closeDialogs();
    let v = await state();
    check(v.header.lot === LOT, S.step, "ロット登録 " + v.header.lot);
    S.step = "work";
    await page.locator("#works button", { hasText: /^H$/ }).first().click(); await W(800); await askAll(page);
    v = await state();
    check(v.header.work !== "", S.step, "作業を選ぶ " + v.header.work);
    await shot(page, "kensa_work");

    // ---- 公差外の確認の窓(選ぶまで閉じない)-------------------------------------
    S.step = "confi";
    // 公差外の値を入れ、その位置(S の厚 1〜4)を測り終えたところで確認が出る(VBA の UFConfi と同じ順)
    for (const t of ["15.400", "15.010", "15.011", "15.012", "15.013"]) {
      if (await page.locator("#confiDlg[open]").count()) break;
      await page.fill("#measIn", `DT10001+00000${t}M`); await page.press("#measIn", "Enter"); await W(600);
    }
    check(await page.locator("#confiDlg[open]").count() === 1, S.step, "公差外で確認の窓が開く");
    await page.keyboard.press("Escape"); await W(300);
    check(await page.locator("#confiDlg[open]").count() === 1, S.step, "Esc では閉じない");
    const rr = await api(page, "POST", "/api/kensa/receive", { text: "DT10001+0000015.020M" });
    check(!!(rr.json && rr.json.stop), S.step, "確認中は測定を受け付けない");
    await page.locator("#confiBox button", { hasText: "許容" }).click(); await W(800); await askAll(page);
    check(await page.locator("#confiDlg[open]").count() === 0, S.step, "許容を選ぶと閉じる");

    // ---- 測定(受信機 → 疑似端末、残りは受信 API)--------------------------------
    S.step = "measure";
    if (PTY !== "-") {
      const cells = (vv) => vv.grid.rows.flatMap((r) => r.cells).filter((c) => c.value === "15.002").length;
      const before = cells(await state());
      await feed("DT10001+0000015.002M"); await W(1500);
      check(cells(await state()) === before + 1, S.step, "受信機(疑似端末)の値が測定表に入る");
      check(await page.locator("#recvWhy").isHidden(), S.step, "入らなかった理由は出ない");
    }
    let done = false;
    for (let i = 0; i < 60 && !done; i++) {
      const t = "DT10001+00000" + (15.003 + (i % 7) * 0.001).toFixed(3) + "M";
      const r = (await api(page, "POST", "/api/kensa/receive", { text: t })).json;
      done = (r.messages || []).some((m) => m.text === "DT検査済");
    }
    check(done, S.step, "厚(DT)を測り終える");
    done = false;
    for (let i = 0; i < 60 && !done; i++) {
      const nx = ((await state()).next.name || "");      // 厚を測り終えた直後は空(#L の 1 つ目で幅へ移る)
      const r = (await api(page, "POST", "/api/kensa/receive", { text: "#L" + (WL[nx.slice(-1)] || "1252.1") })).json;
      done = (r.messages || []).some((m) => /検査済/.test(m.text));
      if (i < 3 || done) note(S.step, `${nx} → ${JSON.stringify((r.messages || []).map((m) => m.text)).slice(0, 160)} ${r.stop ? r.stop.text : ""}`);
    }
    check(done, S.step, "幅・丈(#L)を測り終える");
    // 歪・外観・検入高さは受け渡しの API で入れる(画面の欄の試験は別。印刷前チェックを通すため)
    for (const p of "SUCDE") {
      for (const [n, val] of [[`CB${p}1`, "幅"], [`CB${p}2`, "上"], [`Z${p}`, "0.3"], [`CB${p}3`, "幅"], [`CB${p}4`, "上"],
                              [`G${p}`, "0.2"], [`${p}_UpVisual`, "欠陥ﾅｼ"], [`${p}_UnderVisual`, "欠陥ﾅｼ"]]) {
        const r = await opAll("value", { name: n, value: val });
        if (r.stop) note(S.step, `${n} ${r.stop.text}`);
      }
    }
    await opAll("value", { name: "KenHBox1", value: "200" });
    await page.reload(); await W(1500); await askAll(page); await closeDialogs();

    // ---- 作業範囲の選び欄(F4: 描き直しで作り直さない・選んでいる途中を戻さない)-----------
    S.step = "range";
    await page.click('[data-rtab="mm"]'); await W(200);              // 作業範囲の選び欄は Min/Max の札にある
    const n = await page.locator("#opsSel option").count();
    check(n > 1, S.step, "作業範囲の選択肢 " + n);
    await page.evaluate(() => { document.querySelector("#opsSel option:last-child").dataset.mark = "1"; });
    await page.evaluate(() => document.getElementById("whoToggle").click());     // 描き直し(焦点は動かさない)
    await W(300);
    check(await page.evaluate(() => document.querySelector("#opsSel option:last-child").dataset.mark === "1"),
      S.step, "描き直しで選び欄を作り直さない");
    await page.focus("#opsSel");
    await page.evaluate(() => { document.getElementById("opsSel").value = "3"; });   // 選んでいる途中(change 前)
    await page.evaluate(() => document.getElementById("whoToggle").click());
    await W(300);
    check(await page.evaluate(() => document.getElementById("opsSel").value) === "3", S.step, "選んでいる途中の値を戻さない");
    await page.locator("#measIn").focus();

    // ---- Min/Max ------------------------------------------------------------
    S.step = "minmax";
    await page.selectOption("#opsSel", "1"); await W(400); await askAll(page);
    await page.selectOption("#opeSel", "5"); await W(400); await askAll(page);
    v = await state();
    check(String(v.range.ops) === "1" && String(v.range.ope) === "5", S.step, `範囲 ${v.range.ops}〜${v.range.ope}`);
    await page.click('[data-rtab="ops"]'); await W(200);
    await page.click("#mmBtn"); await W(800); await askAll(page);
    await page.click('[data-rtab="mm"]'); await W(200);
    const mm = await page.locator("#mmKv").innerText();
    check(/15\.\d{3}/.test(mm) && /1252\.\d/.test(mm), S.step, "Min/Max を出す " + mm.replace(/\s+/g, " "));

    // ---- 印刷・保存 -----------------------------------------------------------
    S.step = "print";
    await page.click('[data-rtab="ops"]'); await W(200);
    await page.click("#printBtn"); await W(1500);
    const asked = await askAll(page, "yes", 1500); await W(1500); await askAll(page);
    note(S.step, "聞かれたこと " + JSON.stringify(asked));
    v = await state();
    const ng = (v.checks || []).filter((c) => c.ok === false).map((c) => c.label || c.name || JSON.stringify(c));
    if (ng.length) note(S.step, "印刷前チェックで止まった: " + ng.join(" / "));
    check(!!v.report, S.step, "帳票を作る");
    check(v.buttons.print === "再印刷", S.step, "印刷ボタンが再印刷になる " + v.buttons.print);
    await shot(page, "kensa_printed");

    // ---- サイズ修正の窓 -------------------------------------------------------
    S.step = "resize";
    await page.click("#resizeBtn"); await W(500); await askAll(page); await W(600);
    const rs = await page.evaluate(() => {
      const d = document.querySelector("#tolEdit"); if (!d) return null;
      const r = d.getBoundingClientRect();
      return { visible: r.width > 0 && r.height > 0, sw: d.scrollWidth, cw: d.clientWidth, right: r.right, iw: innerWidth };
    });
    check(rs && rs.visible, S.step, "サイズ修正の窓が開く");
    check(rs && rs.sw <= rs.cw + 1 && rs.right <= rs.iw, S.step, "窓の中が横にはみ出さない " + JSON.stringify(rs));
    await shot(page, "kensa_resize");
    await page.click("#kResizeClose").catch(() => {}); await W(300);

    // ---- 1366×768 で画面がはみ出さない ------------------------------------------
    S.step = "layout";
    await page.reload(); await W(1500); await closeDialogs();
    const sc = await page.evaluate(() => ({ sh: document.documentElement.scrollHeight, ih: innerHeight,
      sw: document.documentElement.scrollWidth, iw: innerWidth }));
    check(sc.sw <= sc.iw, S.step, `横にはみ出さない ${sc.sw}/${sc.iw}`);
    check(sc.sh <= sc.ih, S.step, `縦にはみ出さない ${sc.sh}/${sc.ih}`);
    await shot(page, "kensa_1366");

    // ---- 全ヵ所: 検入高さの欄(F5: 打っている途中の値を描き直しで消さない)-----------------
    S.step = "zenkasho";
    if (PTY !== "-") {
      await page.goto(BASE + "/zenkasho"); await W(1200);
      for (const [op, body] of [["start", { choice_input: "一般" }], ["form", { name: "LotBox", value: "RX021T0" }],
                                ["resume", {}], ["form", { name: "YamaLB", value: "1" }]]) {
        const r = await api(page, "POST", "/api/zenkasho/" + op, body);
        note(S.step, op + " " + r.status + " " + JSON.stringify((r.json && (r.json.stop || r.json.question)) || "").slice(0, 120));
      }
      await page.reload(); await W(1500); await askAll(page);
      const box = page.locator('#kenh input[data-name="KenHBox1"]');
      if (check(await box.count() === 1, S.step, "検入高さの欄 1 が出る")) {
        await box.focus(); await page.keyboard.type("12");             // 打っている途中(change 前)
        await feed("DT10002+0000025.012M"); await W(1800);             // 受信 → 画面の描き直し
        const after = await page.evaluate(() => {
          const i = document.querySelector('#kenh input[data-name="KenHBox1"]');
          return { value: i && i.value, focused: document.activeElement === i };
        });
        check(after.value === "12" && after.focused, S.step, "打っている途中の値と焦点が残る " + JSON.stringify(after));
      }
    }
  } catch (e) {
    check(false, S.step, "例外 " + (e && e.stack || e));
    await shot(page, "fatal_" + S.step).catch(() => {});
  }
  finish(S);
  await browser.close();
})();
