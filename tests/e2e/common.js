// ブラウザの通しの試験の共通部品(tests/e2e/test_e2e.py が node で動かす)。
// 環境変数: E2E_BASE(例 http://127.0.0.1:8877)、E2E_PTY(受信機の代わりの疑似端末。無ければ "-")、E2E_SHOTS(画面の写しの置き場所。任意)
const { chromium } = require("playwright");
const fs = require("fs");

const BASE = process.env.E2E_BASE || "http://127.0.0.1:8877";
const TOKEN = "e2e-token";
const SHOTS = process.env.E2E_SHOTS || "";
const failures = [];
const notes = [];

function note(step, msg) { const l = `[${step}] ${msg}`; notes.push(l); console.log(l); }
function check(cond, step, msg) {
  if (cond) { note(step, "ok   " + msg); return true; }
  failures.push(`[${step}] ${msg}`); console.log(`[${step}] FAIL ${msg}`); return false;
}

async function setup() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 }, deviceScaleFactor: 1,
    extraHTTPHeaders: { "X-Tool-Token": TOKEN } });
  await ctx.addInitScript(() => {        // 読み上げ・音は鳴らさない
    try { window.speechSynthesis.speak = (u) => setTimeout(() => u.onend && u.onend(), 5); } catch (e) { /* 無い */ }
  });
  const page = await ctx.newPage();
  const S = { step: "init", errors: [] };
  page.on("pageerror", (e) => { S.errors.push(`[${S.step}] ${e.message}`); console.log(`[${S.step}] PAGEERROR ${e.message}`); });
  return { browser, ctx, page, S };
}

async function shot(page, name) {
  if (!SHOTS) return;
  fs.mkdirSync(SHOTS, { recursive: true });
  await page.screenshot({ path: `${SHOTS}/${name}.png` });
}

// 開いている「はい/いいえ」に答える(無ければ何もしない)。答えた文を返す
async function askAll(page, answer = "yes", wait = 600) {
  const out = [];
  for (let i = 0; i < 6; i++) {
    const d = page.locator("dialog.askbox[open]");
    try { await d.waitFor({ state: "visible", timeout: wait }); } catch { break; }
    out.push((await d.locator(".askbox__text").textContent()) || "");
    await d.locator(`[data-answer="${answer}"]`).click();
    await page.waitForTimeout(150);
  }
  return out;
}

// 画面と同じ口(トークン付き)で API を呼ぶ
async function api(page, method, url, body) {
  return page.evaluate(async ([m, u, b, t]) => {
    const r = await fetch(u, { method: m, headers: { "Content-Type": "application/json", "X-Tool-Token": t },
      body: m === "GET" ? undefined : JSON.stringify(b || {}) });
    let j = null; try { j = await r.json(); } catch { /* 空 */ }
    return { status: r.status, json: j };
  }, [method, url, body, TOKEN]);
}

// 閉じる。**画面を1つずつ閉じてから**(閉じたこと = pagehide の知らせがサーバに届く)。
// いきなりブラウザを落とすと、次の試験の画面が「ほかのウィンドウで開いています」になる(心拍が途絶えるまで)。
// 閉じた画面は、読み込み直しが引き継ぐあいだ(ufkensa/screens.py の HANDOFF_SEC = 5 秒)使っている画面のままなので、それも待つ
async function close(browser) {
  for (const ctx of browser.contexts()) {
    for (const p of ctx.pages()) await p.close({ runBeforeUnload: false }).catch(() => {});
  }
  await new Promise((r) => setTimeout(r, 5500));
  await browser.close();
}

function finish(S) {
  for (const e of S.errors) failures.push("pageerror " + e);
  console.log("E2E_RESULT " + JSON.stringify({ failures, checks: notes.length }));
  process.exitCode = failures.length ? 1 : 0;
}

module.exports = { BASE, TOKEN, setup, close, shot, askAll, api, note, check, finish };
