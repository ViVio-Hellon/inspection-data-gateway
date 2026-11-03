/*
  設定画面(この端末・使用パス・音・権限・不具合の扱い)。
  判断はサーバが済ませてある(`app/routes/settings.py`)。ここは受け取ったものを並べて、押されたら送るだけ。
*/

import { api } from "../api.js";
import { toast, toastError } from "../toast.js";
import * as tabs from "../tabs.js";
import { previewFile } from "../sound.js";
import { confirmBox } from "../askbox.js";

const $ = (id) => document.getElementById(id);
let S = null;
let SND = null;       // 音の面(`/api/sound/state`)

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
  const canChange = t.unset || S.master_edit.possible;
  fac.disabled = no.disabled = $("saveTerminal").disabled = !canChange;
  $("termLevel").textContent = t.unset ? "未設定" : t.label;
  $("termLevel").className = "lv " + (t.unset ? "lv--ng" : "lv--ok");
  const ln = S.rights.line;
  const notes = [];
  if (ln.text) notes.push(`アクセス権限マスタでは この PC のラインは ${ln.text} です(マスタの値が変わったときだけ自動で当てます)。`);
  else if (ln.problem) notes.push("アクセス権限マスタではラインが決まりません: " + ln.problem);
  if (S.rights.applied) notes.push(S.rights.applied + "。");
  if (!t.unset) notes.push(canChange ? "手で変えるときは管理者パスワードを聞きます。" : "変えるには管理者の権限と管理者パスワードが要ります(権限 の面を見てください)。");
  $("termNote").textContent = notes.join(" ");

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
  input.disabled = !S.master_edit.possible;
  input.setAttribute("aria-label", ariaLabel);
  if (p.key === "log") input.placeholder = "空ならこの端末のローカル(logs)";
  box.append(input);
  const changedKey = field === "path" ? "changed" : field + "_changed";
  const defKey = field === "path" ? "default" : field + "_default";
  if (p[changedKey]) {
    const s = document.createElement("small");
    s.className = "muted";
    s.style.display = "block";
    s.style.overflowWrap = "anywhere";
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
    inp.disabled = !S.master_edit.possible;
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
  const me = S.master_edit;
  const setup = $("pathsSetup");
  setup.hidden = !me.setup;
  setup.textContent = me.setup ? `${me.setup_reason}(設定 → 権限 で確かめられます)。` : "";
  $("pathsLock").textContent = me.possible
    ? (me.ok ? "編集できます(管理者パスワードを入れてあります)。" : "変えるときは管理者パスワードを聞きます。")
    : "変えるには、アクセス権限の「管理者」(Administrator)と管理者パスワードが要ります。";
  const changed = S.paths.filter((p) => p.key !== "log" && (p.single ? p.changed : (p.first_changed || p.second_changed || p.second_names.some((n) => n.value))));
  $("resetAll").disabled = !me.possible || !changed.length;
  $("resetAll").title = changed.length ? `既定から変わっている項目: ${changed.map((p) => p.label).join("・")}` : "すべて既定のままです";
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
    if (p.key === "log" ? S.admin : S.master_edit.possible) {
      const save = document.createElement("button");
      save.className = "btn btn--primary";
      save.type = "button";
      save.textContent = "保存";
      save.addEventListener("click", () => saveRow(p, inputs, names));
      act.append(" ", save);
      const lotChanged = !p.single && (p.first_changed || p.second_changed || p.second_names.some((n) => n.value));
      if (p.key === "log" ? p.changed : (p.single ? p.changed : lotChanged)) {
        const reset = document.createElement("button");
        reset.className = "btn";
        reset.type = "button";
        reset.textContent = "既定に戻す";
        reset.title = p.single ? "既定: " + (p.default || "(空)") : "第1・第2・2つ目の名前をまとめて既定に戻します";
        reset.addEventListener("click", () => (p.key === "log"
          ? postPath({ key: p.key, path: "", reset: true }, "ログの出力パスを既定に戻しました")
          : resetPaths({ key: p.key })));
        act.append(" ", reset);
      }
    }
    tr.append(name, pathTd, status, act);
    return tr;
  }));
}

/* ---- 音(D29)。出来事ごとに「鳴らさない」か音声フォルダのファイル。選ぶとすぐ保存(管理者パスワードは要らない) ---- */
function soundSelect(r) {
  const sel = document.createElement("select");
  sel.className = "input";
  sel.setAttribute("aria-label", `${r.place}・${r.label} の音`);
  const opts = [option(SND.off, SND.off, r.off)];
  const names = [...SND.files];
  if (!r.off && !names.includes(r.file)) opts.push(option(r.file, `${r.file}(フォルダにありません)`, true));
  for (const n of names) opts.push(option(n, n === r.default_file ? `${n}(既定)` : n, !r.off && n === r.file));
  sel.replaceChildren(...opts);
  sel.addEventListener("change", () => saveSound({ key: r.key, file: sel.value }, `${r.place}・${r.label}: ${sel.value}`));
  return sel;
}

function renderSound() {
  if (!SND) return;
  const kv = $("soundDirKv");
  kv.replaceChildren();
  const rows = [["使用パス「音声」", SND.dir || "(空)"],
    ["中の音声ファイル", SND.dir_error ? SND.dir_error : `${SND.files.length} 個`]];
  for (const [k, v] of rows) {
    const dt = document.createElement("dt"); dt.textContent = k;
    const dd = document.createElement("dd"); dd.textContent = v;
    kv.append(dt, dd);
  }
  $("soundDirLv").textContent = SND.dir_error ? "見つかりません" : "読めます";
  $("soundDirLv").className = "lv " + (SND.dir_error ? "lv--ng" : "lv--ok");
  $("soundSuffixes").textContent = SND.suffixes.join(" / ");
  const missing = SND.sounds.filter((r) => !r.off && !r.exists).length;
  const off = SND.sounds.filter((r) => r.off).length;
  $("soundCount").textContent = `${SND.sounds.length} か所 / 鳴らさない ${off} / フォルダにない ${missing}`;
  let lastPlace = "";
  $("soundTable").tBodies[0].replaceChildren(...SND.sounds.map((r) => {
    const tr = document.createElement("tr");
    tr.dataset.key = r.key;
    const place = cell(r.place === lastPlace ? "" : r.place);
    if (r.place !== lastPlace) place.style.fontWeight = "700";
    lastPlace = r.place;
    const name = document.createElement("td");
    const b = document.createElement("b");
    b.textContent = r.label;
    const small = document.createElement("small");
    small.className = "muted";
    small.textContent = ` ${r.key} — VBA ${r.source}`;
    name.append(b, document.createElement("br"), small);
    const when = cell(r.when);
    when.style.whiteSpace = "normal";
    const pick = document.createElement("td");
    pick.append(soundSelect(r));
    if (r.custom) {
      const d = document.createElement("small");
      d.className = "muted";
      d.style.display = "block";
      d.textContent = "既定: " + r.default_file;
      pick.append(d);
    }
    const state = cell(r.off ? "鳴らしません" : (r.exists ? "鳴らせます" : "フォルダにありません"),
      r.off ? "muted" : (r.exists ? "lv--ok" : "lv--ng"));
    state.style.whiteSpace = "normal";
    const act = document.createElement("td");
    const listen = document.createElement("button");
    listen.className = "btn";
    listen.type = "button";
    listen.textContent = "試聴";
    listen.disabled = r.off || !r.exists;
    listen.addEventListener("click", async () => {
      const name = tr.querySelector("select").value;
      if (!(await previewFile(name))) toast(`鳴らせませんでした: ${name}(フォルダにありません)`, "warn");
    });
    act.append(listen);
    if (r.custom) {
      const reset = document.createElement("button");
      reset.className = "btn";
      reset.type = "button";
      reset.textContent = "既定に戻す";
      reset.addEventListener("click", () => saveSound({ key: r.key, reset: true }, `${r.place}・${r.label}: 既定(${r.default_file})に戻しました`));
      act.append(" ", reset);
    }
    tr.append(place, name, when, pick, state, act);
    return tr;
  }));
}

async function loadSound() {
  try {
    SND = await api.get("/api/sound/state");
  } catch (e) {
    SND = null;
    toastError(e);
  }
  renderSound();
  renderBadges();
}

async function saveSound(body, done) {
  try {
    SND = await api.post("/api/settings/sound", body);
    renderSound();
    renderBadges();
    toast("保存しました: " + done, "success");
  } catch (e) {
    toastError(e);
    renderSound();
  }
}

async function checkSoundDir() {
  const out = $("soundCheckResult");
  out.textContent = "確かめています…";
  try {
    const r = await api.post("/api/settings/path/check", { key: SND ? SND.folder_key : "wav" });
    out.textContent = r.message;
    out.className = "path-status " + (r.ok ? "lv--ok" : "lv--ng");
  } catch (e) {
    out.textContent = "";
    toastError(e);
  }
  loadSound();
}

function kvRows(el, rows) {
  el.replaceChildren();
  for (const [k, v, cls] of rows) {
    const dt = document.createElement("dt");
    dt.textContent = k;
    const dd = document.createElement("dd");
    dd.textContent = v;
    if (cls) dd.className = cls;
    el.append(dt, dd);
  }
}

function rowsTable(table, rows) {
  table.tBodies[0].replaceChildren(...rows.map((r) => {
    const tr = document.createElement("tr");
    for (const k of ["行", "管理番号", "ログインID", "PC名", "権限", "有効", "備考"]) tr.append(cell(r[k] == null ? "" : String(r[k])));
    return tr;
  }));
}

function renderRights() {
  const r = S.rights;
  const t = r.table;
  $("rightsLevel").textContent = r.label;
  $("rightsLevel").className = "lv " + (r.admin ? "lv--ok" : "lv--info");
  const setup = $("rightsSetup");
  setup.hidden = !r.setup;
  setup.textContent = r.setup ? S.master_edit.setup_reason + "。使用パス「アクセス権限マスタ」に 梱包資材マスタ.sqlite3(またはそのフォルダ)を入れると、マスタの Administrator の行で決まるようになります。" : "";
  const src = { unset: "未設定", master: "読めました", cache: "届かない(前回の結果)", unreachable: "読めません" }[r.source] || r.source;
  kvRows($("rightsKv"), [
    ["ログインID", r.user || "(不明)", "mono"],
    ["PC名", r.pc || "(不明)", "mono"],
    ["アクセス権限マスタ", t.configured ? t.text : "(未設定)", "mono"],
    ["読むファイル", t.path || `(フォルダを入れたときは その中の ${t.default_file})`, "mono"],
    ["読めたか", t.found ? `${src}(${t.rows} 行)` : `${src}${t.problem ? ": " + t.problem : ""}`, t.found ? "lv--ok" : "lv--ng"],
    ["決まった権限", r.codes.length ? r.codes.join("・") : "なし(一般)"],
    ["この PC のライン", r.line.text || (r.line.problem ? "決めません: " + r.line.problem : "マスタに書いてありません"), r.line.problem ? "lv--ng" : ""],
  ]);
  const notes = [r.reason, ...r.line.notes, r.applied ? r.applied + "。" : ""].filter(Boolean);
  $("rightsReason").textContent = notes.join(" ");
  $("rightsMatchedCount").textContent = `${r.matched.length} 行`;
  rowsTable($("rightsMatched"), r.matched);
  $("rightsIgnored").textContent = r.ignored.length ? "読まなかった権限: " + r.ignored.join("・") : "";
  const codes = [
    ...r.known.map((c) => [c.code, c.label, (c.code === "Administrator" ? r.admin : r.codes.includes(c.code) || r.source === "unset") ? "あり" : "なし"]),
    ...r.aliases.map((a) => [a.code, `${a.same_as} と同じ(前の書き方。互換のため読みます)`, ""]),
    ...r.lines.map((l) => [l.master, `この PC のライン → 設備名 ${l.facility}`, r.line.facility === l.facility ? "この PC" : ""]),
    ["それ以外", "読みません(mode:field・機側・AIM など、ほかのツールの値)", ""],
  ];
  $("codesTable").tBodies[0].replaceChildren(...codes.map(([a, b, c]) => {
    const tr = document.createElement("tr");
    tr.append(cell(a, "mono"), cell(b), cell(c));
    return tr;
  }));
  $("rightsFindCount").textContent = t.found ? `${t.findings.length} 件` : "-";
  $("rightsFindCount").className = "lv " + (t.findings.length ? "lv--warn" : "lv--info");
  const ul = $("rightsFindings");
  ul.replaceChildren(...(t.found ? (t.findings.length ? t.findings : ["気になる行はありません。"]) : ["表を読めていません。"]).map((f) => {
    const li = document.createElement("li");
    li.textContent = f;
    return li;
  }));
  rowsTable($("rightsAdmins"), t.admin_rows);
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
  renderBadges();
}

function renderBadges() {
  if (!S) return;
  const badges = {};
  if (S.terminal.unset) badges.terminal = { text: "未設定", level: "ng" };
  if (S.rights.setup) badges.rights = { text: "設定の途中", level: "warn" };
  else if (!S.admin) badges.rights = { text: "一般", level: "info" };
  else if (S.rights.table.findings.length) badges.rights = { text: `点検 ${S.rights.table.findings.length}`, level: "warn" };
  if (S.master_edit.setup) badges.paths = { text: "設定の途中", level: "warn" };
  if (SND && SND.dir_error) badges.sound = { text: "フォルダなし", level: "warn" };
  else if (SND && SND.sounds.some((r) => !r.off && !r.exists)) badges.sound = { text: "ないファイル", level: "warn" };
  tabs.setBadges($("settingsTabs"), badges);
}

async function load() {
  S = await api.get("/api/settings/state");
  render();
  // 使用パス「音声」を変えたあとも音の面が古くならないよう、設定を読むたびに読み直す
  loadSound();
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
    if (body.key === "wav") loadSound();
    if (done) toast(done, "success");
    return true;
  } catch (e) {
    toastError(e);
    return false;
  }
}

/* 「既定に戻す」(項目ごと)・「すべて既定に戻す」。鍵は使用パスの変更と同じ(403 locked なら窓) */
async function resetPaths(body) {
  try {
    S = await withPathUnlock(() => api.post("/api/settings/path/reset", body));
    render();
    loadSound();
    toast(S.reset ? S.reset.message : "既定に戻しました", "success");
  } catch (e) {
    toastError(e);
  }
}

async function resetAll() {
  const changed = S.paths.filter((p) => p.key !== "log" && (p.single ? p.changed : (p.first_changed || p.second_changed || p.second_names.some((n) => n.value))));
  if (!changed.length) {
    toast("すべて既定のままです", "info");
    return;
  }
  const list = changed.map((p) => "・" + p.label + (p.single ? `: ${p.path || "(空)"} → ${p.default || "(空)"}` : "(第1・第2・2つ目の名前)")).join("\n");
  const ok = await confirmBox(`使用パスを、すべて既定(VBA のオンラインの値)に戻します。\n本番のフォルダを指すようになります。\n\n${list}\n\n(モードとログの出力パスは変えません)`,
    { ok: "すべて既定に戻す", cancel: "やめる", danger: true });
  if (!ok) return;
  await resetPaths({ all: true });
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
  $("soundCheck").addEventListener("click", checkSoundDir);
  $("soundReload").addEventListener("click", loadSound);
  $("soundToPaths").addEventListener("click", (ev) => {
    ev.preventDefault();
    tabs.select($("settingsTabs"), "paths");
    const row = document.querySelector('#pathsTable tr[data-key="wav"]');
    if (row) { row.scrollIntoView({ block: "center" }); const i = row.querySelector("input"); if (i && !i.disabled) i.focus(); }
  });
  $("resetAll").addEventListener("click", resetAll);
  $("rightsToPaths").addEventListener("click", () => {
    tabs.select($("settingsTabs"), "paths");
    const row = document.querySelector('#pathsTable tr[data-key="access_master"]');
    if (row) { row.scrollIntoView({ block: "center" }); const i = row.querySelector("input"); if (i && !i.disabled) i.focus(); }
  });
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
