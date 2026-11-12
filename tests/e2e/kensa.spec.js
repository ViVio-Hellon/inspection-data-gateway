// 検査画面の通しの試験(2026-11-10 リファクタリング 第1段 5)。tests/e2e/test_e2e.py から node で動かす。
// 受信機の状態 → 検査員・ロット・作業 → 公差外の確認の窓(選ぶまで閉じない)→ 受信機(疑似端末)と受信 API で測定
// → 作業範囲の選び欄(F4)→ Min/Max → 印刷・保存 → サイズ修正の窓 → 1366×768 ではみ出さない → 全ヵ所の検入高さ(F5)
const { BASE, setup, close, shot, askAll, api, note, check, finish } = require("./common");

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
    await W(300);
    check(await page.evaluate(() => document.activeElement && document.activeElement.id === "measIn"), S.step,
      "測定に入ると受け取る欄(見えない)に焦点がある");
    // 2026-11-11: 不良要因は 1 つだけ選ぶ(別の行で入れ替え、同じ行で外す)
    await page.click('[data-tab="bad"]'); await W(300);
    const drows = page.locator('#defectList [role="option"]');
    if (await drows.count() >= 2) {
      await drows.nth(0).click(); await drows.nth(1).click();
      const sel = () => page.locator('#defectList [role="option"][aria-selected="true"]').count();
      check(await sel() === 1 && await drows.nth(1).getAttribute("aria-selected") === "true", S.step, "不良要因は 1 つだけ選ぶ");
      await shot(page, "kensa_bad_tab");
      await drows.nth(1).click();
      check(await sel() === 0, S.step, "選んだ行をもう一度押すと外す");
    }
    await page.click('[data-tab="strain"]'); await W(200);

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
    // 2026-11-11: 許容の印を押すと「許容を外しますか」(ここでは「いいえ」)
    const allowB = page.locator('#mgrid .ngflag--allow[data-pos="S"]').first();
    check(await allowB.count() === 1, S.step, "許容の印は押せる");
    if (await allowB.count()) {
      await allowB.click();
      const asked = await askAll(page, "no");
      check(asked.some((t) => /許容を外しますか/.test(t)), S.step, "許容を外すか聞く");
    }

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
    // 右の列の「操作」の札は、列の中でもスクロールせずに収まる(印刷の後 = 印刷履歴クリア・最近の操作が出ている状態)
    const opsFit = async (label) => {
      await page.click('[data-rtab="ops"]'); await W(250);
      const f = await page.evaluate(() => {
        const k = document.querySelector("#krail"), kb = k.getBoundingClientRect();
        const btns = [...document.querySelectorAll('[data-rpane="ops"] .btn')].filter((b) => b.offsetParent);
        const low = Math.max(...btns.map((b) => b.getBoundingClientRect().bottom));
        return { sh: k.scrollHeight, ch: k.clientHeight, low: Math.round(low), bottom: Math.round(Math.min(kb.bottom, innerHeight)),
                 n: btns.length, page: document.documentElement.scrollHeight, ih: innerHeight };
      });
      check(f.sh <= f.ch + 1 && f.low <= f.bottom && f.page <= f.ih, S.step, `${label}: 操作の札が列の中でスクロールせずに収まる ${JSON.stringify(f)}`);
    };
    await opsFit("1366×768 印刷の後");
    await page.reload(); await W(1500); await closeDialogs();
    const sc = await page.evaluate(() => ({ sh: document.documentElement.scrollHeight, ih: innerHeight,
      sw: document.documentElement.scrollWidth, iw: innerWidth }));
    check(sc.sw <= sc.iw, S.step, `横にはみ出さない ${sc.sw}/${sc.iw}`);
    check(sc.sh <= sc.ih, S.step, `縦にはみ出さない ${sc.sh}/${sc.ih}`);
    await shot(page, "kensa_1366");
    // 明暗の配色 × 1366×768 / 1920×1080 でも同じ(写しは kensa_ops_<幅>x<高さ>_<配色>.png)
    for (const theme of ["light", "dark"]) {
      for (const [w, h] of [[1366, 768], [1920, 1080]]) {
        await page.setViewportSize({ width: w, height: h });
        await page.evaluate((t) => { document.documentElement.dataset.theme = t; }, theme); await W(400);
        await opsFit(`${w}×${h} ${theme}`);
        await shot(page, `kensa_ops_${w}x${h}_${theme}`);
      }
    }
    await page.evaluate(() => { delete document.documentElement.dataset.theme; });
    await page.setViewportSize({ width: 1366, height: 768 }); await W(400);

    // ---- 2026-11-11 現場の声: 歪・不良・検入高さ・BH・包装の知らせ(1366×768 暗い配色)-----------------
    S.step = "tabs";
    await page.evaluate(() => { document.documentElement.dataset.theme = "dark"; }); await W(300);
    await page.click('[data-tab="strain"]'); await W(300);
    // 2026-11-11: 右の列に受け取る欄(テキストボックス)を見せない。状態の 1 行と「次に測る」は出す
    const rin = await page.evaluate(() => { const r = document.getElementById("measIn").getBoundingClientRect(); return r.right; });
    check(rin <= 0 && await page.locator("#nextBox").isVisible() && await page.locator("#recvState").isVisible(), S.step,
      "受け取る欄は画面の外・状態と次に測るは見える " + rin);
    await shot(page, "kensa_right_dark");
    const tol = await page.evaluate(() => [...document.querySelectorAll("#strainTol .strain-tol__item")].map((x) => ({
      k: x.dataset.k, border: getComputedStyle(x).borderLeftWidth, fs: parseFloat(getComputedStyle(x).fontSize) })));
    check(tol.length === 5 && tol.every((t) => t.border === "6px" && t.fs >= 18), S.step, "歪の公差を大きく色の枠で出す " + JSON.stringify(tol));
    const zh = await page.evaluate(() => { const i = document.querySelector('#strainGrid input[data-name^="Z"]'); return i ? i.getBoundingClientRect().height : 0; });
    check(zh >= 38, S.step, "歪の入力欄を大きく " + zh);
    await shot(page, "kensa_strain_dark");
    check(await page.locator('[data-tab="bh"]').isHidden(), S.step, "BH の札は隠す");
    // 不良: 枚数を打って Enter を押さずに「反映」→「枚数入力がありません」にならない(バグ 3)
    await page.click('[data-tab="bad"]'); await W(300);
    const lab = await page.evaluate(() => [...document.querySelectorAll('#defectList [role="option"]')].slice(0, 3).map((o) => {
      const n = o.querySelector(".deflist__name"), c = o.querySelector(".deflist__code"), r = o.getBoundingClientRect();
      return { name: n.textContent, code: c.textContent, codeRight: Math.round(r.right - c.getBoundingClientRect().right) };
    }));
    check(lab[0].name === "フクレ" && lab[0].code === "01" && lab[1].name === "" && lab[1].code === "02" && lab[2].name === "ムシレ"
      && lab.every((x) => x.codeRight === lab[0].codeRight), S.step, "不良要因は 名前 左・コード 右 " + JSON.stringify(lab));
    await page.click('#badGrid input[data-name="SBad"]'); await W(400);
    await page.locator('#badGrid input[data-name="SBad"]').fill("2");
    await page.click('#defectList [data-value="0"]');
    await shot(page, "kensa_bad_dark");
    const respP = page.waitForResponse((r) => r.url().endsWith("/api/kensa/defect-reflect"));
    await page.click("#reflectBtn");
    const rj = await (await respP).json();
    const texts = (rj.messages || []).map((m) => m.text);
    check(!texts.includes("枚数入力がありません") && !rj.stop, S.step, "打った枚数で反映できる " + JSON.stringify(texts));
    check(((rj.view.bad || []).find((c) => c.pos === "S") || {}).value === "", S.step, "反映のあと枚数を空に戻す");
    // 不良要因票の見本: 反映した枚数が入力の横の見本に出る(印刷と同じ帳票。2026-11-11)
    await W(800);
    const pv = await page.evaluate(() => {
      const f = document.getElementById("defPrev"), b = document.getElementById("defPrevBox").getBoundingClientRect();
      const d = f && !f.hidden && f.contentDocument;
      const tds = d ? [...d.querySelectorAll("td")].map((t) => t.textContent.trim()) : [];
      return { shown: !!d, has2: tds.includes("2"), right: Math.round(b.right), bottom: Math.round(b.bottom),
               sw: d ? d.documentElement.scrollWidth : 0, fw: f ? f.clientWidth : 0 };
    });
    check(pv.shown && pv.has2, S.step, "反映すると不良要因票の見本に枚数が出る " + JSON.stringify(pv));
    check(pv.right <= 1366 && pv.bottom <= 768 && pv.sw <= pv.fw + 2, S.step, "見本は画面と箱に収まる " + JSON.stringify(pv));
    await shot(page, "kensa_bad_preview_dark");
    await page.evaluate(() => document.documentElement.setAttribute("data-theme", "light")); await W(200);
    await shot(page, "kensa_bad_preview_light");
    await page.evaluate(() => document.documentElement.setAttribute("data-theme", "dark")); await W(200);
    // 検入高さ: 厚板(H)では検査回数を出さない(少枚数だけ)
    await page.click('[data-tab="kenh"]'); await W(300);
    check(v.header.work !== "少枚数" && (await page.locator('#kenhBoxes input[data-name="KenCount"]').count()) === 0,
      S.step, "少枚数でなければ検査回数を出さない " + v.header.work);
    await shot(page, "kensa_kenh_H_dark");
    // 包装仕様の注意・梱包保護材: 1 つの大きな窓。確認を押すまで閉じない
    await page.evaluate(async () => {
      const jb = await import("/static/js/judgebox.js");
      jb.collectPacking({ kind: "packing_spec", text: ",B,直角度の測定をすること,先行データ票が必要です。" });
      jb.collectPacking({ kind: "packing_material", text: "アングル" });
      jb.collectPacking({ kind: "packing_material", text: "上部保護材:発泡シート" });
    });
    await W(400);
    const secs = await page.evaluate(() => [...document.querySelectorAll("#packingBox[open] .packbox__sec h3")].map((h) => h.textContent));
    check(JSON.stringify(secs) === JSON.stringify(["包装仕様の注意", "梱包保護材"]), S.step, "包装の知らせを 1 つの窓に " + JSON.stringify(secs));
    await shot(page, "kensa_packing_dark");
    await page.keyboard.press("Escape"); await W(200);
    check(await page.locator("#packingBox[open]").count() === 1, S.step, "Esc では閉じない");
    await page.click("#packingBox .packbox__ok"); await W(200);
    check(await page.locator("#packingBox[open]").count() === 0, S.step, "確認で閉じる");
    await page.evaluate(() => { delete document.documentElement.dataset.theme; });

    // ---- 2026-11-12 現場の声: 検査員・ロットは別の入力欄 / 歪は欄を離れても判定 / 表示の選び方 / ラテラルボー要 ----
    S.step = "add-split";
    await page.setViewportSize({ width: 1366, height: 768 }); await W(300);
    await page.locator("#ksteps li").nth(0).click(); await W(300);
    check(await page.locator("#kAdd[open]").count() === 1 && await page.locator("#cardIn").isVisible()
      && await page.locator("#lotIn").isHidden(), S.step, "検査員の追加は社員カードの欄だけ");
    await page.locator("#cardIn").fill("12345"); await page.keyboard.press("Escape"); await W(300);
    await page.locator("#ksteps li").nth(2).click(); await W(300);
    check(await page.locator("#lotIn").isVisible() && await page.locator("#cardIn").isHidden()
      && await page.locator("#lotIn").inputValue() === "", S.step, "ロットの読み直しは LotNo の欄だけ・空");
    await page.locator("#lotIn").fill("H73"); await page.click("#kAddClose"); await W(300);
    await page.locator("#ksteps li").nth(0).click(); await W(300);
    check(await page.locator("#cardIn").inputValue() === "", S.step, "やめたあと開き直すと打ちかけは残らない");
    await page.click("#kAddClose"); await W(300);

    S.step = "layout-mode";
    const fits = async () => page.evaluate(() => ({ sh: document.documentElement.scrollHeight, ih: innerHeight,
      sw: document.documentElement.scrollWidth, iw: innerWidth }));
    await page.click('[data-rtab="ops"]'); await W(200);
    for (const [w, h] of [[1920, 1080], [1366, 768]]) {
      await page.setViewportSize({ width: w, height: h }); await W(300);
      await page.click("#layoutToggle"); await W(400);
      const m = await page.evaluate(() => ({ all: document.getElementById("kensa").classList.contains("kensa--alltab"),
        ctab: document.getElementById("kensa").dataset.ctab, gridTab: !!document.querySelector(".subtab--grid").offsetParent,
        sub: !!document.getElementById("subCard").offsetParent, mg: !!document.getElementById("gridCard").offsetParent }));
      check(m.all && m.ctab === "grid" && m.gridTab && m.mg, S.step, `${w}×${h} すべてタブ: 測定表がタブの 1 つ ` + JSON.stringify(m));
      let f = await fits();
      check(f.sh <= f.ih && f.sw <= f.iw, S.step, `${w}×${h} すべてタブでページがスクロールしない ` + JSON.stringify(f));
      await shot(page, `kensa_layout_alltab_grid_${w}x${h}`);
      await page.click('[data-tab="strain"]'); await W(300);
      check(await page.locator("#gridCard").isHidden() && await page.locator("#strainGrid").isVisible(), S.step, `${w}×${h} 歪のタブで測定表は隠れる`);
      f = await fits();
      check(f.sh <= f.ih && f.sw <= f.iw, S.step, `${w}×${h} 歪タブでもスクロールしない ` + JSON.stringify(f));
      await shot(page, `kensa_layout_alltab_strain_${w}x${h}`);
      if (PTY !== "-") {
        await feed("DT10001+0000015.010M"); await W(1500);
        check(await page.evaluate(() => document.getElementById("kensa").dataset.ctab) === "grid", S.step, `${w}×${h} 測定値が来たら測定表のタブへ`);
      }
      await page.click("#layoutToggle"); await W(400);
      const back = await page.evaluate(() => document.getElementById("kensa").classList.contains("kensa--alltab"));
      check(!back, S.step, `${w}×${h} 測定+タブに戻せる`);
      f = await fits();
      check(f.sh <= f.ih && f.sw <= f.iw, S.step, `${w}×${h} 測定+タブでスクロールしない ` + JSON.stringify(f));
      await shot(page, `kensa_layout_split_${w}x${h}`);
    }
    check(await page.evaluate(() => localStorage.getItem("kensa.layout")) === "split", S.step, "選んだ表示を覚える");

    S.step = "strain-blur";
    await page.click('[data-tab="strain"]'); await W(300);
    const sent = [];
    const onReq = (r) => { if (r.url().endsWith("/api/kensa/value")) { try { sent.push(JSON.parse(r.postData() || "{}").name); } catch { /* */ } } };
    page.on("request", onReq);
    const z = page.locator('#strainGrid input[data-name^="Z"]').first();
    const zname = await z.getAttribute("data-name");
    await z.click(); await z.fill("0.3"); await page.keyboard.press("Tab"); await W(800);
    check(sent.filter((n) => n === zname).length === 1, S.step, "Enter なしで欄を離れると 1 回送る " + JSON.stringify(sent));
    sent.length = 0;
    const z2 = page.locator(`#strainGrid input[data-name="${zname}"]`);
    await z2.click(); await z2.fill("0.4"); await page.keyboard.press("Enter"); await W(600);
    await page.locator('[data-tab="visual"]').click(); await W(400);
    check(sent.filter((n) => n === zname).length === 1, S.step, "Enter で送った値は離れても送り直さない " + JSON.stringify(sent));
    page.off("request", onReq);
    await page.click('[data-tab="strain"]'); await W(200);

    S.step = "lateral-hit";
    // 当たりのロットは手元の台帳に無いので、画面へ渡る view に当たりを足して描かせる(表示だけの確かめ)
    const inject = async (route) => {
      const r = await route.fetch(); const j = await r.json();
      if (j && j.view) {
        j.view.lateral_hit = { on: true, who: "ナメカワ", text: "3.0", source: ["製品規格のラテラルボー要", "DIM_179 日車/白銅滑川の表"] };
        (j.view.grid.lateral || []).forEach((c) => { if (c.pos === "S" || c.pos === "U") c.visible = true; });
      }
      await route.fulfill({ response: r, json: j });
    };
    await page.route("**/api/kensa/state", inject);
    await page.route("**/api/kensa/start", inject);
    await page.reload(); await W(1800); await closeDialogs();
    await page.click('[data-rtab="tol"]'); await W(300);
    const lh = await page.locator("#latHit").innerText().catch(() => "");
    check(/ラテラルボー要: ナメカワ/.test(lh) && /DIM_179/.test(lh), S.step, "公差のタブにラテラルボー要を出す " + lh);
    const rowHit = await page.locator("#mgrid .lathit--row").innerText().catch(() => "");
    check(/ナメカワ/.test(rowHit), S.step, "測定表のラテラルボーの行にも出す " + rowHit);
    await shot(page, "kensa_lateral_hit_1366x768");
    await page.unroute("**/api/kensa/state"); await page.unroute("**/api/kensa/start");
    await page.click('[data-rtab="ops"]'); await W(200);

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
  await close(browser);
})();
