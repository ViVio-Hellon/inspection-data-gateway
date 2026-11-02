/*
  設定画面(この端末・使用パス・権限・不具合の扱い)。
  判断はサーバが済ませてある(`app/routes/settings.py`)。ここは受け取ったものを並べて、押されたら送るだけ。
*/

import { api } from "../api.js";
import { toast, toastError } from "../toast.js";
import * as tabs from "../tabs.js";

const $ = (id) => document.getElementById(id);
let S = null;
let viewMode = null;

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

function renderPaths() {
  viewMode = viewMode || S.mode;
  const pick = $("pathsModePick");
  pick.replaceChildren(...S.modes.map((m) => option(m.key, m.label + (m.key === S.mode ? "(いま使っている)" : ""), m.key === viewMode)));
  $("pathsMode").textContent = "いま: " + (S.modes.find((m) => m.key === S.mode) || {}).label;
  const body = $("pathsTable").tBodies[0];
  body.replaceChildren(...S.paths[viewMode].map((p) => {
    const tr = document.createElement("tr");
    tr.dataset.key = p.key;
    const name = document.createElement("td");
    name.style.whiteSpace = "normal";
    name.style.width = "38%";
    const b = document.createElement("b");
    b.textContent = p.label;
    const small = document.createElement("small");
    small.className = "muted";
    small.textContent = ` ${p.vba} — ${p.use}`;
    name.append(b, document.createElement("br"), small);
    const pathTd = document.createElement("td");
    const input = document.createElement("input");
    input.className = "input";
    input.style.width = "100%";
    input.style.minWidth = "22em";
    input.value = p.path;
    input.disabled = !S.admin;
    input.setAttribute("aria-label", p.label + " のパス");
    if (p.key === "log") input.placeholder = "空ならこの端末のローカル(logs)";
    pathTd.append(input);
    if (p.changed) {
      const s = document.createElement("small");
      s.className = "muted";
      s.textContent = "既定: " + (p.default || (p.key === "log" ? "この端末のローカル(logs)" : "(空)"));
      pathTd.append(document.createElement("br"), s);
    }
    const status = cell("", "path-status");
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
      save.addEventListener("click", () => savePath(p.key, input.value));
      act.append(" ", save);
      if (p.changed) {
        const reset = document.createElement("button");
        reset.className = "btn";
        reset.type = "button";
        reset.textContent = "既定に戻す";
        reset.addEventListener("click", () => savePath(p.key, "", true));
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
    viewMode = S.mode;
    render();
    toast("モードを変えました", "success");
  } catch (e) {
    toastError(e);
    render();
  }
}

async function savePath(key, path, reset = false) {
  try {
    S = await api.post("/api/settings/path", { key, path, mode: viewMode, reset });
    render();
    toast(reset ? "既定に戻しました" : "保存しました", "success");
  } catch (e) {
    toastError(e);
  }
}

async function checkOne(key, td) {
  td.textContent = "確かめています…";
  try {
    const r = await api.post("/api/settings/path/check", { key, mode: viewMode });
    td.textContent = r.message;
    td.className = "path-status " + (r.ok ? "lv--ok" : "lv--ng");
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
  $("pathsModePick").addEventListener("change", () => { viewMode = $("pathsModePick").value; renderPaths(); });
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
