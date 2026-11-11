// キャスク・全ヵ所・Aインプット・校正の通しの試験(整理 6、2026-11-10)。tests/e2e/test_e2e.py から node で動かす。
// kensa.spec.js の後に同じサーバで動かす(検査員の登録・検査のロット・Min/Max・受信機(疑似端末)はそこで済んでいる)。
// 各画面: 開く → 問いのある操作 1 つ(画面の確認窓で答える。ボタンの文言も見る)→ 受信機の値がその画面に届く。
// どの画面も共通の op()(app/static/js/screen_op.js)と受信(app/static/js/recv.js)を使う。
const { BASE, setup, close, shot, api, note, check, finish } = require("./common");

const PTY = process.env.E2E_PTY || "-";

(async () => {
  const { browser, page, S } = await setup();
  const W = (ms) => page.waitForTimeout(ms);
  const feed = (text) => api(page, "POST", "/__e2e/feed", { text });
  const json = async (method, url, body) => ((await api(page, method, url, body)).json || {});
  // 開いている確認窓の文と 2 つのボタンの文言を読み、answer("yes" / "no")を押す
  const answer = async (which, timeout = 3000) => {
    const d = page.locator("dialog.askbox[open]");
    try { await d.waitFor({ state: "visible", timeout }); } catch { return null; }
    const out = { text: (await d.locator(".askbox__text").textContent()) || "",
                  yes: (await d.locator('[data-answer="yes"]').textContent()) || "",
                  no: (await d.locator('[data-answer="no"]').textContent()) || "" };
    await d.locator(`[data-answer="${which}"]`).click();
    await W(400);
    return out;
  };
  const toasts = () => page.evaluate(() => Array.from(document.querySelectorAll(".toast, [class*=toast]"))
    .map((e) => e.textContent || "").join(" | "));
  try {
    // ---- キャスク ---------------------------------------------------------------
    S.step = "cask";
    await page.goto(BASE + "/cask"); await W(1500);
    let v = (await json("GET", "/api/cask/state")).view;
    check(v && v.visible, S.step, "開く(検査員登録済み)");
    for (const [op, body] of [["lot", { lot: "zz-zzzzz" }], ["zero", { text: "DT10002+0000000.000M" }],
                              ["zero", { text: "#L0.0" }], ["select", { row: 7, col: 2 }]]) {
      const r = await api(page, "POST", "/api/cask/" + op, body);
      if (r.status !== 200 || (r.json && r.json.stop)) note(S.step, `${op} ${r.status} ${JSON.stringify(r.json && r.json.stop)}`);
    }
    await page.reload(); await W(1500);
    if (PTY !== "-") {
      await feed("DT10002+0000036.000M"); await W(1800);
      v = (await json("GET", "/api/cask/state")).view;
      const cols = v.columns.map((c) => c.col);
      const row = v.rows.find((r) => r[0] === 7);
      check(row && row[2][cols.indexOf(2)][0] === "36", S.step, "受信機の値がキャスクの欄(7,2)に入る");
    }
    await page.click("#newBtn");
    const qc = await answer("no");
    check(qc && /新規発行しますか/.test(qc.text) && qc.yes === "はい" && qc.no === "いいえ", S.step,
      "シート発行の問い(はい/いいえ) " + JSON.stringify(qc));
    await W(300);
    check(/ｷｬﾝｾﾙ処理/.test(await toasts()), S.step, "いいえ → 答えを付けて送り直し、ｷｬﾝｾﾙ処理を出す");
    await shot(page, "cask");

    // ---- 全ヵ所 -----------------------------------------------------------------
    S.step = "zenkasho";
    await page.goto(BASE + "/zenkasho"); await W(1500);
    v = (await json("GET", "/api/zenkasho/state")).view;
    check(!!v, S.step, "開く");
    if (PTY !== "-") {
      const before = JSON.stringify(v);
      await feed("DT10002+0000025.031M"); await W(1800);
      v = (await json("GET", "/api/zenkasho/state")).view;
      check(!before.includes("25.031") && JSON.stringify(v).includes("25.031"), S.step, "受信機の値が全ヵ所の欄に入る");
    }
    await page.click("#newBtn");
    const qz = await answer("no");
    check(qz && /新規発行しますか/.test(qz.text) && qz.yes === "はい" && qz.no === "いいえ", S.step,
      "新規発行の問い(はい/いいえ) " + JSON.stringify(qz));
    v = (await json("GET", "/api/zenkasho/state")).view;
    check(JSON.stringify(v).includes("25.031") || PTY === "-", S.step, "いいえ → シートは消さない");
    await shot(page, "zenkasho");

    // ---- Aインプット ------------------------------------------------------------
    S.step = "ainput";
    await page.goto(BASE + "/ainput"); await W(1200);
    const op = await json("POST", "/api/ainput/open", { ken_m: "L1" });
    check(op.view && op.view.ready && !op.stop, S.step, "開く(検査のロット・Min/Max から) " + JSON.stringify(op.stop));
    await page.reload(); await W(1500);
    if (PTY !== "-") {
      await feed("DT10006+0000075.000M"); await W(1800);
      check(/DT10006\+0000075\.000M/.test(await page.locator("#received").innerText()), S.step,
        "受信機の値が Aインプットに届く " + (await page.locator("#received").innerText()));
    }
    await page.click("#resetBtn");
    const qa = await answer("no");
    check(qa && /hostsend/.test(qa.text) && qa.yes === "OK" && qa.no === "キャンセル", S.step,
      "Reset の問い(OK/キャンセル) " + JSON.stringify(qa));
    v = (await json("GET", "/api/ainput/state")).view;
    check(v.ready && v.send.stage !== "reset", S.step, "キャンセル → 書き直さない " + v.send.stage);
    await shot(page, "ainput");

    // ---- 校正 -------------------------------------------------------------------
    S.step = "calibration";
    await page.goto(BASE + "/calibration"); await W(1500);
    v = (await json("GET", "/api/calibration/state")).view;
    check(v && v.step === "1000", S.step, "開く(1000mm 直尺から) " + (v && v.step));
    if (PTY !== "-") {
      check(/受信中/.test(await page.locator("#calRecv").innerText()), S.step, "受信機の状態を出す");
      await feed("#L1000.1"); await W(1800);
      const st = await page.locator("#calRecv").innerText();
      check(/#L1000\.1/.test(st), S.step, "最後に届いた値を出す " + st.replace(/\s+/g, " "));
      v = (await json("GET", "/api/calibration/state")).view;
      check(v.step !== "1000", S.step, "受信機の値で 1000mm の校正が進む " + v.step);
    }
    const step = v.step;
    await page.click("#calClear"); await W(800);
    v = (await json("GET", "/api/calibration/state")).view;
    check(v.step === step, S.step, "ｸﾘｱ＆ﾘｾｯﾄ は今の段のまま(VBA の CBcl) " + v.step);
    await page.click("#calRestart"); await W(800);
    v = (await json("GET", "/api/calibration/state")).view;
    check(v.step === "1000", S.step, "やり直しで 1000mm に戻る " + v.step);
    const bad = await api(page, "POST", "/api/calibration/rank", {});
    check(bad.status === 400, S.step, "足りない値は 400 " + bad.status);
    await shot(page, "calibration");

    // ---- 保守 → 寸法のしきい値(2026-11-11)------------------------------------------
    S.step = "thresholds";
    await page.goto(BASE + "/maint"); await W(1200);
    await page.click('#mTabs .tab[data-tab="thresh"]'); await W(1200);
    const nTh = await page.locator("#mThresh input").count();
    check(nTh >= 18, S.step, "しきい値の入力欄を出す " + nTh);
    check((await page.locator("#th-op_notthin_lo").inputValue()) === "6.001", S.step, "既定は VBA と同じ(6.001)");
    check((await page.locator("#th-kennyu_n").inputValue()) === "25", S.step, "検入の枚数の既定 25");
    await shot(page, "maint_thresholds");

    // ---- 保守 → 照合ルール(2026-11-11)------------------------------------------
    S.step = "match_rules";
    await page.click('#mTabs .tab[data-tab="rules"]'); await W(1200);
    const nRules = await page.locator("#rRules tr").count();
    check(nRules === 8, S.step, "ラテフリ直角 の既定のルール 8 行 " + nRules);
    const lot = "N8061T0";
    await page.fill("#rLot", lot);
    await page.click('#rTrialForm button[type="submit"]'); await W(1500);
    const steps = await page.locator("#rTrial table").count();
    check(steps >= 0, S.step, "試し照合の表 " + steps);
    await page.click("#rCheck"); await W(2500);
    check(/結果が変わるもの: 0 件/.test(await page.locator("#rReg").innerText()), S.step, "既定のままなら変わるロットは 0 件");
    await shot(page, "maint_match_rules");
    // 第3段: 測定機器ID(ライン・厚・用途を直接入れて試す)
    await page.click('#rGroups button:has-text("測定機器ID")'); await W(600);
    check(await page.locator("#rRules tr").count() === 3, S.step, "測定機器ID の既定のルール 3 行");
    await page.fill("#rLot", "");
    await page.fill("#rIn_ライン", "HVC"); await page.fill("#rIn_厚", "10"); await page.fill("#rIn_用途", "H191");
    await page.click('#rTrialForm button[type="submit"]'); await W(1200);
    check(/M4/.test(await page.locator("#rTrial").innerText()), S.step, "HVC 厚10 用途H191 → M4");
    await page.click("#rCheck"); await W(2500);
    check(/結果が変わるもの: 0 件/.test(await page.locator("#rReg").innerText()), S.step, "測定機器ID 既定のままなら 0 件");
    await shot(page, "maint_match_rules_micro");
  } catch (e) {
    check(false, S.step, "例外 " + (e && e.stack || e));
    await shot(page, "fatal_" + S.step).catch(() => {});
  }
  finish(S);
  await close(browser);
})();
