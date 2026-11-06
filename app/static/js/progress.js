/* 少しでも待つ操作の進み具合(2026-11-06 のご依頼: ロット検索・検査員検索で重さを感じる)。

   画面の要求が 300ms を過ぎても返らないとき、画面の下に「いま何をしているか(段の名前)・経過・進み具合」を出す。
   段の名前と進み具合はサーバの /api/progress(trace.progress)から取る。進み具合は前回の同じ操作の段の数から
   見積もるので、初めての操作は段の名前と経過だけ(横棒は動き続ける)。カード・LotNo の自動送信にも効く。 */
const QUIET = /^\/api\/(progress|health|jobs|kensa\/recv|maint\/rx|log)/;
// 操作の名前(サーバの一番外の処理名 → 画面に出す言葉)。無いものは段の名前だけ出す
const TITLES = {
  "kensa.lot": "ロットを読み込み中", "kensa.worker": "検査員を確かめ中", "kensa.start": "検査画面を準備中",
  "kensa.work": "作業選択を反映中", "kensa.print-check": "印刷前チェック中", "kensa.print-commit": "印刷・保存中",
  "kensa.new": "新規検査の準備中", "kensa.minmax": "Min/Max を計算中", "kensa.lot-source": "控えから読み直し中",
  "cask.lot": "ロットを読み込み中", "zenkasho.lot": "ロットを読み込み中", "quality.lot": "品質基準を読み込み中",
};
let pending = 0;
let timer = null;
let poll = null;
let box = null;
let t0 = 0;

function ensureBox() {
  if (box && document.body.contains(box)) return box;
  box = document.createElement("div");
  box.className = "progress-pop";
  box.setAttribute("role", "status");
  box.setAttribute("aria-live", "polite");
  box.innerHTML = '<div class="progress-pop__row"><b class="progress-pop__label">処理中…</b>'
    + '<span class="progress-pop__sec"></span></div><div class="progress-pop__track"><i></i></div>';
  box.hidden = true;
  document.body.append(box);
  return box;
}

async function refresh(token) {
  try {
    const res = await fetch("/api/progress", { cache: "no-store", headers: { "X-Tool-Token": token } });
    const d = await res.json();
    const p = (d.running || [])[0];
    const b = ensureBox();
    const sec = ((performance.now() - t0) / 1000).toFixed(1);
    const title = p && TITLES[p.op];
    b.querySelector(".progress-pop__label").textContent = p ? (title ? `${title} — ${p.label}` : p.label) : "処理中…";
    b.querySelector(".progress-pop__sec").textContent = `${sec}秒` + (p && p.pct != null ? ` ・ ${p.pct}%` : "");
    const bar = b.querySelector(".progress-pop__track i");
    if (p && p.pct != null) { bar.style.width = `${p.pct}%`; bar.classList.remove("is-indef"); }
    else { bar.style.width = "35%"; bar.classList.add("is-indef"); }
  } catch (e) { /* 次の回で */ }
}

/** 要求の始まり。返り値の関数を終わりで呼ぶ。 */
export function begin(path, token) {
  if (QUIET.test(path)) return () => {};
  pending += 1;
  if (pending === 1) {
    t0 = performance.now();
    timer = setTimeout(() => {
      const b = ensureBox();
      b.hidden = false;
      refresh(token);
      poll = setInterval(() => refresh(token), 250);
    }, 300);
  }
  let ended = false;
  return () => {
    if (ended) return;
    ended = true;
    pending = Math.max(0, pending - 1);
    if (pending === 0) {
      clearTimeout(timer);
      clearInterval(poll);
      if (box) box.hidden = true;
    }
  };
}
