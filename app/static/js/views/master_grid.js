/*
  マスタの表 1 つを見る・直す部品(保守の「マスタ管理」と「測定機器ID」で使う)。
  python-web-tools のマスタ管理(master.js)と同じ形: 絞り込み(全部の列)・見出しで並び替え・行を押すと 1 行の窓
  (直す・消す)・行を足す・変更履歴。直せるか・値の形・行がほかで変わっていないかはサーバ(master_admin.py)が決める。
*/

import { api } from "../api.js";
import { confirmBox } from "../askbox.js";
import { withUnlock } from "../admin_unlock.js";
import { toast, toastError } from "../toast.js";

const ROW_KEY = "__rowid";

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text ?? "";
  return e;
}

function show(v) { return v === null || v === undefined ? "" : String(v); }

export function createMasterGrid(root, { compact = false, onChanged = null } = {}) {
  let view = null;        // サーバが返した最後のページ
  let src = "", table = "", sort = "", dir = "asc";
  let editing = null;     // 開いている行(足すときは null)

  root.classList.add("mgrid");
  const head = el("div", "mgrid__head");
  const title = el("b", "mgrid__title", "表を選んでください");
  const count = el("span", "muted mgrid__count");
  const form = el("form", "mgrid__find");
  form.autocomplete = "off";
  const q = el("input", "input mgrid__q");
  q.placeholder = "絞り込み(どの列でも)";
  q.type = "search";
  const findBtn = el("button", "btn btn--sm", "絞る");
  findBtn.type = "submit";
  form.append(q, findBtn);
  const addBtn = el("button", "btn btn--sm btn--select", "＋ 行を足す");
  const histBtn = el("button", "btn btn--sm", "変更履歴");
  const reloadBtn = el("button", "btn btn--sm", "読み直す");
  for (const b of [addBtn, histBtn, reloadBtn]) b.type = "button";
  head.append(title, count, form, addBtn, histBtn, reloadBtn);
  const why = el("p", "mgrid__why");
  const wrap = el("div", "tablewrap mgrid__wrap" + (compact ? " mgrid__wrap--compact" : ""));
  const tbl = el("table", "table table--dense mgrid__table");
  wrap.append(tbl);
  root.replaceChildren(head, why, wrap);

  // ---- 1 行の窓 ----
  const dlg = el("dialog", "mgrid__dlg");
  dlg.innerHTML = `<form method="dialog" class="mgrid__dform" autocomplete="off">
      <div class="mgrid__dhead"><b class="mgrid__dtitle"></b><span class="muted mgrid__dsub"></span></div>
      <p class="mgrid__dwhy"></p>
      <div class="mgrid__fields"></div>
      <p class="mgrid__derr"></p>
      <div class="mgrid__dfoot">
        <button class="btn btn--danger" type="button" data-act="delete">この行を消す</button>
        <span class="mgrid__gap"></span>
        <button class="btn" type="button" data-act="close">閉じる</button>
        <button class="btn btn--commit" type="submit" data-act="save">保存</button>
      </div></form>`;
  document.body.append(dlg);
  const $d = (s) => dlg.querySelector(s);

  // ---- 変更履歴の窓 ----
  const hdlg = el("dialog", "mgrid__dlg mgrid__hdlg");
  hdlg.innerHTML = `<div class="mgrid__dform"><div class="mgrid__dhead"><b>変更履歴</b><span class="muted mgrid__hsub"></span></div>
      <div class="tablewrap mgrid__hwrap"><table class="table table--dense"><thead><tr><th>日時</th><th>誰が</th><th>表</th><th>操作</th><th>行</th><th>前 → 後</th></tr></thead><tbody></tbody></table></div>
      <div class="mgrid__dfoot"><span class="mgrid__gap"></span><button class="btn" type="button">閉じる</button></div></div>`;
  document.body.append(hdlg);
  hdlg.querySelector("button").addEventListener("click", () => hdlg.close());

  let mark_fn = null, scroll_mark = false;
  function render() {
    if (!view) return;
    title.textContent = `${view.label} / ${view.table}`;
    count.textContent = view.query ? `${view.total} 件(絞り込み中)` : `${view.total} 件`;
    addBtn.hidden = !view.editable;
    const notes = [];
    if (!view.editable) notes.push("見るだけ: " + (view.why || "この DB は直せません。"));
    else notes.push("行を押すと 1 行の窓が開きます(直す・消す)。保存には編集の鍵が要ります。");
    if (view.note) notes.push(view.note);
    why.textContent = notes.join("\n");
    why.classList.toggle("is-ro", !view.editable);

    const thead = el("thead");
    const hr = el("tr");
    for (const c of view.columns) {
      const th = el("th");
      const b = el("button", "mgrid__sort", c.name);
      b.type = "button";
      b.title = `${c.kind_label}${c.required ? "・必須" : ""}${c.pk ? "・主キー" : ""}(押すと並び替え)`;
      if (view.sort === c.name) b.append(el("span", "mgrid__arrow", view.sort_dir === "desc" ? " ▼" : " ▲"));
      b.addEventListener("click", () => {
        if (sort === c.name) dir = dir === "asc" ? "desc" : "asc"; else { sort = c.name; dir = "asc"; }
        load();
      });
      th.append(b);
      hr.append(th);
    }
    thead.append(hr);
    const tbody = el("tbody");
    if (!view.rows.length) {
      const tr = el("tr");
      const td = el("td", "muted", view.query ? "絞り込みに合う行がありません。" : "行がありません。");
      td.colSpan = view.columns.length || 1;
      tr.append(td);
      tbody.append(tr);
    }
    let first = null;
    for (const r of view.rows) {
      const mark = mark_fn ? mark_fn(r) : null;
      const tr = el("tr", "mgrid__row" + (mark ? ` is-${mark}` : ""));
      if (mark && !first) first = tr;
      tr.tabIndex = 0;
      for (const c of view.columns) {
        const v = show(r[c.name]);
        const td = el("td", c.kind === "text" ? "" : "n", v);
        if (v.length > 24) td.title = v;
        tr.append(td);
      }
      tr.addEventListener("click", () => openRow(r));
      tr.addEventListener("keydown", (e) => { if (e.key === "Enter") openRow(r); });
      tbody.append(tr);
    }
    tbl.replaceChildren(thead, tbody);
    if (first && scroll_mark) { first.scrollIntoView({ block: "nearest" }); scroll_mark = false; }
  }

  async function load() {
    if (!src || !table) return;
    try {
      const p = new URLSearchParams({ source: src, table, q: q.value.trim(), sort, dir });
      view = await api.get(`/api/maint/master/rows?${p}`);
      render();
    } catch (err) {
      view = null;
      title.textContent = table;
      count.textContent = "";
      why.textContent = err.message;
      why.classList.add("is-ro");
      tbl.replaceChildren();
      addBtn.hidden = true;
    }
  }

  function openRow(row) {
    editing = row;
    const ro = !view.editable;
    $d(".mgrid__dtitle").textContent = row ? `${view.table} の 1 行` : `${view.table} に行を足す`;
    $d(".mgrid__dsub").textContent = row ? `(行番号 ${row[ROW_KEY]})` : "";
    $d(".mgrid__dwhy").textContent = ro ? "見るだけ: " + (view.why || "") : "＊ は必須。空欄は「無し」(NULL)になります。";
    $d(".mgrid__derr").textContent = "";
    const fields = $d(".mgrid__fields");
    fields.replaceChildren(...view.columns.flatMap((c) => {
      const id = "mgf_" + Math.random().toString(36).slice(2);
      const lab = el("label", "", c.name + (c.required ? " ＊" : ""));
      lab.htmlFor = id;
      const kind = el("span", "mgrid__kind", c.auto ? "自動の番号" : c.kind_label);
      const long = row && show(row[c.name]).length > 40;
      const inp = el(long ? "textarea" : "input", "input" + (c.kind === "text" ? "" : " input--num"));
      inp.id = id;
      inp.dataset.col = c.name;
      inp.value = row ? show(row[c.name]) : "";
      if (long) inp.rows = 3;
      inp.readOnly = ro;
      if (c.auto && !row) inp.placeholder = "(空なら自動)";
      return [lab, inp, kind];
    }));
    $d('[data-act="delete"]').hidden = ro || !row;
    $d('[data-act="save"]').hidden = ro;
    $d('[data-act="save"]').textContent = row ? "保存" : "足す";
    dlg.showModal();
    fields.querySelector("input:not([readonly]), textarea:not([readonly])")?.focus();
  }

  function seen(row) {
    const s = { ...row };
    delete s[ROW_KEY];
    return s;
  }

  async function write(body, done) {
    $d(".mgrid__derr").textContent = "";
    for (const x of dlg.querySelectorAll(".is-bad")) x.classList.remove("is-bad");
    try {
      const d = await withUnlock(() => api.post("/api/maint/master/row", { source: src, table, ...body }));
      if (!d) return;
      dlg.close();
      toast(d.message || done, "ok");
      await load();
      onChanged?.({ source: src, table });
    } catch (err) {
      $d(".mgrid__derr").textContent = err.message;
      const t = err.field && dlg.querySelector(`[data-col="${CSS.escape(err.field)}"]`);
      if (t) { t.classList.add("is-bad"); t.focus(); }
      if (err.code === "stale") load();
    }
  }

  $d("form").addEventListener("submit", (e) => {
    e.preventDefault();
    const values = {};
    for (const inp of dlg.querySelectorAll("[data-col]")) {
      const c = inp.dataset.col;
      if (!editing || inp.value !== show(editing[c])) values[c] = inp.value;
    }
    if (editing) {
      if (!Object.keys(values).length) { $d(".mgrid__derr").textContent = "変わった値がありません。"; return; }
      write({ action: "save", rowid: editing[ROW_KEY], values, seen: seen(editing) }, "直しました");
    } else {
      write({ action: "add", values }, "足しました");
    }
  });
  $d('[data-act="close"]').addEventListener("click", () => dlg.close());
  $d('[data-act="delete"]').addEventListener("click", async () => {
    if (!editing) return;
    if (!(await confirmBox(`${view.table} のこの行(行番号 ${editing[ROW_KEY]})を消します。元に戻せません。よろしいですか？`,
                           { ok: "消す", cancel: "やめる", danger: true }))) return;
    write({ action: "delete", rowid: editing[ROW_KEY], seen: seen(editing) }, "消しました");
  });

  form.addEventListener("submit", (e) => { e.preventDefault(); load(); });
  q.addEventListener("search", () => load());
  addBtn.addEventListener("click", () => openRow(null));
  reloadBtn.addEventListener("click", () => load());
  histBtn.addEventListener("click", async () => {
    try {
      const d = await api.get(`/api/maint/master/history?${new URLSearchParams({ source: src, table })}`);
      hdlg.querySelector(".mgrid__hsub").textContent = `${view ? view.label + " / " + view.table : ""}(直近 30 日・この PC の記録)`;
      const fmt = (o) => (o ? Object.entries(o).map(([k, v]) => `${k}=${show(v)}`).join(", ") : "");
      const rows = d.rows.map((h) => {
        const tr = el("tr");
        const change = h.action === "直す"
          ? Object.keys(h.after || {}).map((k) => `${k}: ${show((h.before || {})[k]) || "(空)"} → ${show(h.after[k]) || "(空)"}`).join("\n")
          : h.action === "消す" ? "消した行: " + fmt(h.before) : "足した行: " + fmt(h.after);
        tr.append(el("td", "n", (h.t || "").replace("T", " ").slice(0, 19)), el("td", "", h.user || h.pc),
                  el("td", "", h.table), el("td", "", h.action), el("td", "n", show(h.rowid)), el("td", "mgrid__change", change));
        return tr;
      });
      if (!rows.length) { const tr = el("tr"); const td = el("td", "muted", "まだ変更はありません。"); td.colSpan = 6; tr.append(td); rows.push(tr); }
      hdlg.querySelector("tbody").replaceChildren(...rows);
      hdlg.showModal();
    } catch (err) { toastError(err); }
  });

  return {
    /** 行に印を付ける(fn(行) → "ok" / "ng" / "expect" / null)。測定機器ID 画面で当たった行を光らせる */
    highlight(fn) { mark_fn = fn; scroll_mark = true; render(); },
    open(source, tableName) {
      if (source !== src || tableName !== table) { sort = ""; dir = "asc"; q.value = ""; }
      src = source; table = tableName;
      return load();
    },
    reload: load,
  };
}
