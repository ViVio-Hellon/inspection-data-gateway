/*
  設定画面(この端末・使用パス・権限・不具合の扱い)。
  判断はサーバが済ませてある(`app/routes/settings.py`)。ここは受け取ったものを並べて、押されたら送るだけ。
*/

import { api } from "../api.js";
import { toast, toastError } from "../toast.js";
import * as tabs from "../tabs.js";

const $ = (id) => document.getElementById(id);
let S = null;

function option(value, label, selected) {
  const o = document.createElement("option");
  o.value = value;
  o.textContent = label;
  o.selected = selected;
  return o;
}

function cell(text, cls) {
  const td = document.createElement("td");
  td.textContent = text;
  if (cls) td.className = cls;
  return td;
}

function renderTerminal() {
  const t = S.terminal;
  const fac = $("facility");
  fac.replaceChildren(option("", "(選んでください)", !t.facility),
    ...t.facilities.map((f) => option(f, f, f === t.facility)));
  const no = $("lineNo");
  no.replaceChildren(option("", "-", !t.line_no), ...t.numbers.map((n) => option(n, n, n === t.line_no)));
  const maru = fac.value === "丸徳";
  no.hidden = !maru;
  $("lineNoLabel").hidden = !maru;
  const canChange = t.unset || S.admin;
  fac.disabled = no.disabled = $("saveTerminal").disabled = !canChange;
  $("termLevel").textContent = t.unset ? "未設定" : t.label;
  $("termLevel").className = "lv " + (t.unset ? "lv--ng" : "lv--ok");
  $("termNote").textContent = canChange ? "" : "変えるには管理者の権限が要ります(権限 の面を見てください)。";

  const box = $("modeChoices");
  box.replaceChildren(...S.modes.map((m) => {
    const label = document.createElement("label");
    const input = document.createElement("input");
    input.type = "radio";
    input.name = "mode";
    input.value = m.key;
    input.checked = m.key === S.mode;
    input.disabled = !S.admin;
    input.addEventListener("change", () => saveMode(m.key));
    label.append(input, " ", m.label);
    return label;
  }));
}

function pathInput(p, field, ariaLabel) {
  const box = document.createElement("div");
  const input = document.createElement("input");
  input.className = "input";
  input.style.width = "100%";
  input.style.minWidth = "22em";
  input.value = p[field] ?? "";
  input.dataset.field = field;
  input.disabled = !S.admin;
  input.setAttribute("aria-label", ariaLabel);
  if (p.key === "log") input.placeholder = "空ならこの端末のローカル(logs)";
  box.append(input);
  const changedKey = field === "path" ? "changed" : field + "_changed";
  const defKey = field === "path" ? "default" : field + "_default";
  if (p[changedKey]) {
    const s = document.createElement("small");
    s.className = "muted";
    s.textContent = "既定: " + (p[defKey] || (p.key === "log" ? "この端末のローカル(logs)" : "(空)"));
    box.append(s);
  }
  return { box, input };
}

function lotPathCell(p, td, inputs, names) {
  for (const [field, label] of [["first", "第1パス(台帳 = MainPath)"], ["second", "第2パス(控え = NextPath)"]]) {
    const head = document.createElement("small");
    head.textContent = label;
    head.style.fontWeight = "700";
    const f = pathInput(p, field, `仕掛台帳の${label}`);
    td.append(head, f.box);
    inputs.push(f.input);
  }
  const box = document.createElement("div");
  box.className = "muted";
  box.style.marginTop = "4px";
  const t = document.createElement("small");
  t.textContent = "2つ目の名前(任意。空なら1つ目と同じ名前):";
  box.append(t);
  for (const n of p.second_names) {
    const line = document.createElement("div");
    const lab = document.createElement("small");
    lab.textContent = n.first + " → ";
    const inp = document.createElement("input");
    inp.className = "input";
    inp.style.width = "14em";
    inp.value = n.value;
    inp.placeholder = n.first;
    inp.disabled = !S.admin;
    inp.dataset.db = n.db;
    inp.dataset.orig = n.value;
    inp.setAttribute("aria-label", `2つ目の名前(${n.first} の代わり)`);
    line.append(lab, inp);
    box.append(line);
    names.push(inp);
  }
  td.append(box);
  const use = document.createElement("div");
  use.style.marginTop = "4px";
  for (const m of S.modes) {
    const d = document.createElement("small");
    d.style.display = "block";
    d.textContent = `${m.label}: ${p.in_use[m.key]}`;
    if (m.key === S.mode) d.style.fontWeight = "700";
    use.append(d);
  }
  td.append(use);
}

function renderPaths() {
  const modeLabel = (S.modes.find((m) => m.key === S.mode) || {}).label;
  $("pathsMode").textContent = "いま: " + modeLabel;
  $("pathsLock").textContent = S.admin
    ? (S.master_edit.ok ? "編集できます(管理者パスワードを入れてあります)。" : "変えるときは管理者パスワードを聞きます。")
    : "変えるには、アクセス権限の「管理者」と管理者パスワードが要ります。";
  const body = $("pathsTable").tBodies[0];
  body.replaceChildren(...S.paths.map((p) => {
    const tr = document.createElement("tr");
    tr.dataset.key = p.key;
    const name = document.createElement("td");
    name.style.whiteSpace = "normal";
    name.style.width = "30%";
    const b = document.createElement("b");
    b.textContent = p.label;
    const small = document.createElement("small");
    small.className = "muted";
    small.textContent = ` ${p.vba} — ${p.use}`;
    name.append(b, document.createElement("br"), small);
    const pathTd = document.createElement("td");
    const inputs = [];
    const names = [];
    if (p.single) {
      const f = pathInput(p, "path", p.label + " のパス");
      pathTd.append(f.box);
      inputs.push(f.input);
    } else {
      lotPathCell(p, pathTd, inputs, names);
    }
    const status = cell("", "path-status");
    status.style.whiteSpace = "normal";
    const act = document.createElement("td");
    const check = document.createElement("button");
    check.className = "btn";
    check.type = "button";
    check.textContent = "確認";
    check.addEventListener("click", () => checkOne(p.key, status));
    act.append(check);
    if (S.admin) {
      const save = document.createElement("button");
      save.className = "btn btn--primary";
      save.type = "button";
      save.textContent = "保存";
      save.addEventListener("click", () => saveRow(p, inputs, names));
      act.append(" ", save);
      const resets = p.single ? (p.changed ? [["", "既定に戻す"]] : [])
        : ["first", "second"].filter((f) => p[f + "_changed"]).map((f) => [f, (f === "first" ? "第1" : "第2") + "を既定に"]);
      for (const [slot, text] of resets) {
        const reset = document.createElement("button");
        reset.className = "btn";
        reset.type = "button";
        reset.textContent = text;
        reset.addEventListener("click", () => postPath({ key: p.key, slot, path: "", reset: true }, "既定に戻しました"));
        act.append(" ", reset);
      }
    }
    tr.append(name, pathTd, status, act);
    return tr;
  }));
}

function renderRights() {
  const r = S.rights;
  $("rightsLevel").textContent = r.label;
  $("rightsLevel").className = "lv " + (S.admin ? "lv--ok" : "lv--info");
  const kv = $("rightsKv");
  kv.replaceChildren();
  for (const [k, v] of [["ログインID", r.user], ["PC名", r.pc], ["読んだ表", r.table.path || "(なし)"]]) {
    const dt = document.createElement("dt");
    dt.textContent = k;
    const dd = document.createElement("dd");
    dd.textContent = v;
    kv.append(dt, dd);
  }
  $("rightsReason").textContent = r.reason;
  $("codesTable").tBodies[0].replaceChildren(...r.known.map((c) => {
    const tr = document.createElement("tr");
    const on = c.code === "ufkensa:admin" ? S.admin : r.codes.includes(c.code) || r.source === "unset";
    tr.append(cell(c.code, "mono"), cell(c.label), cell(on ? "あり" : "なし"));
    return tr;
  }));
  $("rightsTable").textContent = r.table.problem ? "表を読めませんでした: " + r.table.problem : "";
}

function renderQuirks() {
  const fix = new Set(S.quirks.fix);
  $("quirksCount").textContent = `直す ${fix.size} 件 / 再現 ${Object.keys(S.quirks.titles).length - fix.size} 件`;
  $("quirksTable").tBodies[0].replaceChildren(...Object.entries(S.quirks.titles).map(([k, title]) => {
    const tr = document.createElement("tr");
    tr.append(cell(k, "mono"), cell(title), cell(fix.has(k) ? "直す" : "VBA どおり再現"));
    return tr;
  }));
}

function renderLog() {
  const L = S.log;
  $("logNames").textContent = `${L.tool_id}_${L.pc}_日付.log と ${L.tool_id}_記録_${L.pc}_日付.jsonl`;
  $("logLv").textContent = L.problem ? "書けません" : "書けます";
  $("logLv").className = "lv " + (L.problem ? "lv--ng" : "lv--ok");
  const kvEl = $("logKv");
  kvEl.replaceChildren();
  for (const [k, v] of [["いまの置き場所", L.dir], ["この端末のローカル", L.default], ["書けない理由", L.problem || "-"]]) {
    const dt = document.createElement("dt"); dt.textContent = k;
    const dd = document.createElement("dd"); dd.textContent = v;
    kvEl.append(dt, dd);
  }
  if (document.activeElement !== $("logDir")) $("logDir").value = L.requested;
  if (document.activeElement !== $("logKeep")) $("logKeep").value = L.keep_days;
  $("logDir").disabled = $("logKeep").disabled = $("saveLog").disabled = !S.admin;
  $("logCount").textContent = `${L.problems.length} 件`;
  $("logTable").tBodies[0].replaceChildren(...L.problems.map((p) => {
    const tr = document.createElement("tr");
    tr.append(cell(p.t, "mono"), cell(p.lv), cell(p.proc || p.ev), cell(p.msg), cell(p.ms == null ? "" : `${p.ms} ms`, "mono"));
    return tr;
  }));
}

function render() {
  renderTerminal();
  renderPaths();
  renderRights();
  renderQuirks();
  renderLog();
  const root = $("settingsTabs");
  const badges = {};
  if (S.terminal.unset) badges.terminal = { text: "未設定", level: "ng" };
  if (!S.admin) badges.rights = { text: "一般", level: "info" };
  tabs.setBadges(root, badges);
}

async function load() {
  S = await api.get("/api/settings/state");
  render();
}

async function saveTerminal() {
  try {
    S = await api.post("/api/settings/terminal", { facility: $("facility").value, line_no: $("lineNo").value });
    render();
    toast("この端末のラインを決めました: " + S.terminal.label, "success");
  } catch (e) {
    toastError(e);
  }
}

async function saveMode(mode) {
  try {
    S = await api.post("/api/settings/mode", { mode });
    render();
    toast("モードを変えました", "success");
  } catch (e) {
    toastError(e);
    render();
  }
}

/* 使用パスの変更は管理者パスワードが要る(D26)。403 code "locked" なら鍵の窓(admin_unlock.js)を出して 1 回だけ呼び直す */
async function withPathUnlock(fn) {
  try {
    return await fn();
  } catch (e) {
    if (e.status !== 403 || e.code !== "locked") throw e;
    let mod = null;
    try {
      mod = await import("../admin_unlock.js");
    } catch (_) {
      mod = null;
    }
    if (!mod || typeof mod.ensureUnlocked !== "function") throw e;
    if (!(await mod.ensureUnlocked(e.message))) throw e;
    return await fn();
  }
}

async function postPath(body, done) {
  try {
    S = await withPathUnlock(() => api.post("/api/settings/path", body));
    render();
    if (done) toast(done, "success");
    return true;
  } catch (e) {
    toastError(e);
    return false;
  }
}

async function saveRow(p, inputs, names) {
  const bodies = [];
  for (const input of inputs) {
    const field = input.dataset.field;
    if (input.value === (p[field] ?? "")) continue;
    bodies.push(field === "path" ? { key: p.key, path: input.value } : { key: p.key, slot: field, path: input.value });
  }
  for (const n of names) {
    if (n.value !== n.dataset.orig) bodies.push({ key: p.key, db: n.dataset.db, second_name: n.value });
  }
  if (!bodies.length) {
    toast("変わったところはありません", "info");
    return;
  }
  for (const b of bodies) {
    if (!(await postPath(b, ""))) return;
  }
  toast("保存しました", "success");
}

async function checkOne(key, td) {
  td.textContent = "確かめています…";
  try {
    const r = await api.post("/api/settings/path/check", { key });
    td.replaceChildren();
    for (const x of r.results) {
      const d = document.createElement("div");
      d.textContent = (x.label ? x.label.replace("パス", "") + ": " : "") + x.message;
      d.className = x.ok ? "lv--ok" : "lv--ng";
      td.append(d);
    }
    if (r.in_use) {
      const d = document.createElement("small");
      d.className = "muted";
      d.textContent = "いま読む側: " + r.in_use;
      td.append(d);
    }
    td.className = "path-status";
  } catch (e) {
    td.textContent = "";
    toastError(e);
  }
}

function init() {
  tabs.attach($("settingsTabs"));
  $("facility").addEventListener("change", () => {
    const maru = $("facility").value === "丸徳";
    $("lineNo").hidden = !maru;
    $("lineNoLabel").hidden = !maru;
  });
  $("saveTerminal").addEventListener("click", saveTerminal);
  $("checkAll").addEventListener("click", () => {
    for (const tr of $("pathsTable").tBodies[0].rows) checkOne(tr.dataset.key, tr.querySelector(".path-status"));
  });
  $("saveLog").addEventListener("click", async () => {
    try {
      S = await api.post("/api/settings/log", { dir: $("logDir").value, keep_days: $("logKeep").value });
      render();
      toast("ログの設定を保存しました", "success");
    } catch (e) {
      toastError(e);
    }
  });
  $("reloadRights").addEventListener("click", async () => {
    try {
      S = await api.post("/api/settings/rights/reload", {});
      render();
      toast("権限を読み直しました", "success");
    } catch (e) {
      toastError(e);
    }
  });
  load().catch(toastError);
}

init();
