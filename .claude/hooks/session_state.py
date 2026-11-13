"""SESSION_STATE.md を圧縮の前後で守る hook(標準ライブラリのみ・Windows/Linux 共通)。

  python session_state.py precompact   … PreCompact(auto / manual)
  python session_state.py sessionstart … SessionStart(startup / resume / compact / clear)

precompact:
  - .claude/SESSION_STATE.md が無ければ雛形を作る
  - 控えを .claude/state/ に置く(新しい 5 個だけ残す)
  - Git の状況(ブランチ・最後のコミット・変更ファイル)を SESSION_STATE.md の自動欄に書き直す
    (毎回置き換えるので、たまっていかない)
sessionstart:
  - SESSION_STATE.md の中身を、作業を続けるための文脈としてモデルに渡す

どちらも何が起きても exit 0。圧縮・起動を止めない。
"""
from __future__ import annotations

import datetime as _dt
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

KEEP_BACKUPS = 5
MAX_CONTEXT_CHARS = 12000
AUTO_BEGIN = "<!-- AUTO:GIT BEGIN(圧縮の直前に hook が書き直す。手で直さない) -->"
AUTO_END = "<!-- AUTO:GIT END -->"

TEMPLATE = """# SESSION_STATE

圧縮(compact)で会話が要約されても作業を続けられるように、今の状態を短く書いておく。
完了した項目は「完了した作業」に1行で残し、細部は消す(ためこまない)。

## 現在の作業目的
-

## 完了した作業
-

## 未完了の作業
-

## 重要な仕様・設計上の決定
-

## 発生した不具合と解決状況
-

## 変更したファイル
-

## 次に実行する作業
-

{begin}
{end}
""".format(begin=AUTO_BEGIN, end=AUTO_END)


def _read_stdin() -> dict:
    try:
        raw = sys.stdin.read()
        return json.loads(raw) if raw.strip() else {}
    except Exception:  # noqa: BLE001
        return {}


def _project_dir(payload: dict) -> Path:
    d = os.environ.get("CLAUDE_PROJECT_DIR") or payload.get("cwd") or os.getcwd()
    p = Path(d).resolve()
    # 下のフォルダから呼ばれても、.claude のあるところまで上がる
    for q in (p, *p.parents):
        if (q / ".claude").is_dir():
            return q
    return p


def _git(root: Path, *args: str) -> str:
    try:
        r = subprocess.run(["git", *args], cwd=str(root), capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=10)
        return r.stdout.strip() if r.returncode == 0 else ""
    except Exception:  # noqa: BLE001
        return ""


def _git_block(root: Path, trigger: str) -> str:
    now = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [AUTO_BEGIN, f"### Git の状況(自動 {now} / 圧縮 {trigger or '-'})"]
    branch = _git(root, "rev-parse", "--abbrev-ref", "HEAD")
    if not branch:
        lines += ["- Git の情報は取れませんでした", AUTO_END]
        return "\n".join(lines)
    lines.append(f"- ブランチ: `{branch}`")
    last = _git(root, "log", "-1", "--format=%h %s")
    if last:
        lines.append(f"- 最後のコミット: {last}")
    status = _git(root, "status", "--short")
    files = [s for s in status.splitlines() if s.strip()]
    if files:
        lines.append(f"- まだコミットしていない変更: {len(files)} 件")
        lines += [f"  - `{s}`" for s in files[:30]]
        if len(files) > 30:
            lines.append(f"  - …ほか {len(files) - 30} 件")
    else:
        lines.append("- まだコミットしていない変更: なし")
    ahead = _git(root, "rev-list", "--count", "@{u}..HEAD")
    if ahead and ahead != "0":
        lines.append(f"- まだ push していないコミット: {ahead} 件")
    lines.append(AUTO_END)
    return "\n".join(lines)


def _replace_auto(text: str, block: str) -> str:
    b, e = text.find(AUTO_BEGIN), text.find(AUTO_END)
    if b != -1 and e != -1 and e > b:
        return text[:b] + block + text[e + len(AUTO_END):]
    return text.rstrip() + "\n\n" + block + "\n"


def _write_atomic(path: Path, text: str) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8", newline="\n")
    os.replace(tmp, path)


def precompact(payload: dict) -> dict:
    root = _project_dir(payload)
    state = root / ".claude" / "SESSION_STATE.md"
    keep = root / ".claude" / "state"
    keep.mkdir(parents=True, exist_ok=True)
    if not state.exists():
        _write_atomic(state, TEMPLATE)
    stamp = _dt.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    shutil.copy2(state, keep / f"SESSION_STATE.{stamp}.md")
    for old in sorted(keep.glob("SESSION_STATE.*.md"))[:-KEEP_BACKUPS]:
        try:
            old.unlink()
        except OSError:
            pass
    trigger = str(payload.get("trigger") or "")
    text = state.read_text(encoding="utf-8", errors="replace")
    _write_atomic(state, _replace_auto(text, _git_block(root, trigger)))
    return {"systemMessage": f"SESSION_STATE.md を控えました(.claude/state/SESSION_STATE.{stamp}.md)"}


def sessionstart(payload: dict) -> dict:
    root = _project_dir(payload)
    state = root / ".claude" / "SESSION_STATE.md"
    if not state.exists():
        return {}
    text = state.read_text(encoding="utf-8", errors="replace")
    if len(text) > MAX_CONTEXT_CHARS:
        text = text[:MAX_CONTEXT_CHARS] + "\n…(長いので途中まで。全文は .claude/SESSION_STATE.md)"
    src = str(payload.get("source") or "")
    head = ("圧縮のあとです。" if src == "compact" else "") + \
        ("以下は .claude/SESSION_STATE.md(作業の引き継ぎ)。作業目的・未完了の作業・次に実行する作業を確かめてから続けること。"
         "作業が進んだら、このファイルを更新する(完了した項目は1行にまとめ、ためこまない)。\n\n")
    return {"hookSpecificOutput": {"hookEventName": "SessionStart", "additionalContext": head + text}}


def main() -> int:
    mode = sys.argv[1] if len(sys.argv) > 1 else ""
    payload = _read_stdin()
    try:
        out = precompact(payload) if mode == "precompact" else sessionstart(payload) if mode == "sessionstart" else {}
    except Exception as exc:  # noqa: BLE001 - 圧縮・起動を止めない
        out = {"systemMessage": f"SESSION_STATE の hook で問題がありました(続けます): {exc}"}
    if out:
        sys.stdout.write(json.dumps(out, ensure_ascii=True))   # Windows の cp932 でも化けない
    return 0


if __name__ == "__main__":
    sys.exit(main())
