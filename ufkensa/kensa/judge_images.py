"""判定の画像(VBA の UFOK / UFNG / UF対角 フォームの絵。2026-11-09)

VBA は絵をフォームに埋め込み、判定のたびに出していた:
- UFOK / UFNG(UFOK.frm:15-26 / UFNG.frm:16-25): 合格・不合格で `Show vbModeless`、3 秒で閉じる
  (全ヵ所検査 ZenSh2.cls:78-100・Cask.bas:1429-1450。音楽合格 / 音楽不合格 と同時)
- UF対角(UF対角.frm:12-50): 対角線の欄に値が入ったとき(Cask.bas:1899 UF対角起動。音楽対角 と同時)、
  DB1〜DB5 のうち 1 枚(検査シートは行 14〜18 で決まり、ほかは乱数)を 2 秒出す

ここでは、使用パス「判定の画像」のフォルダに決まった名前で置いた画像を配る。
2026-11-09: VBA のフォームに埋め込まれていた絵(UF検査判定.xlsm の vbaProject.bin から取り出したもの)を
`assets/judge/` に同梱し、使用パスが空・その画像が無いときはそれを出す(設定しなくても VBA と同じ絵が出る)。
名前: OK.* / NG.* / 対角1.*〜対角5.*(DB1〜DB5。1 枚だけなら 対角.*)。拡張子は jpg・jpeg・png・gif・bmp。
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

EXT = (".jpg", ".jpeg", ".png", ".gif", ".bmp")
KINDS = {"ok": ("OK",), "ng": ("NG",), "diag": ("対角1", "対角2", "対角3", "対角4", "対角5", "対角")}
LABELS = {"ok": "OK", "ng": "NG", "diag": "対角線"}
SHOW_MS = {"ok": 3000, "ng": 3000, "diag": 2000}       # UFOK・UFNG は 3 秒、UF対角 は 2 秒


#: 同梱の絵(VBA の UFOK / UFNG / UF対角 から取り出したもの)
DEFAULT_DIR = Path(__file__).resolve().parent / "assets" / "judge"


def root() -> Optional[Path]:
    from . import paths
    text = paths.get("judge_images")
    if not text:
        return None
    return Path(paths.to_local(text))


def _find(folder: Path, stem: str) -> Optional[Path]:
    try:
        for f in folder.iterdir():
            if f.is_file() and f.suffix.lower() in EXT and f.stem.upper() == stem.upper():
                return f
    except OSError:
        return None
    return None


def files(kind: str, folder: Optional[Path] = None) -> list[Path]:
    """その種類の画像(あるものだけ)。フォルダが空・無いときは []。"""
    folder = root() if folder is None else folder
    if folder is None or kind not in KINDS:
        return []
    try:
        if not folder.is_dir():
            return []
    except OSError:
        return []
    out = [f for f in (_find(folder, s) for s in KINDS[kind]) if f is not None]
    return out


def effective(kind: str) -> list[Path]:
    """出す画像。使用パスのフォルダにあればそれ、無ければ同梱の絵。"""
    return files(kind) or files(kind, DEFAULT_DIR)


def available(folder: Optional[Path] = None) -> dict[str, list[str]]:
    """{種類: [ファイル名…]}(画面が出せる画像だけ)。"""
    if folder is None:
        return {k: [f.name for f in effective(k)] for k in KINDS}
    return {k: [f.name for f in files(k, folder)] for k in KINDS}


def find(kind: str, name: str) -> Optional[Path]:
    """配ってよいファイル(その種類の決まった名前のものだけ)。"""
    for f in effective(kind):
        if f.name == name:
            return f
    return None


def listing(folder: Path) -> list[dict]:
    """設定画面の確認用。"""
    out = []
    for k in KINDS:
        fs = files(k, folder)
        out.append({"kind": k, "label": LABELS[k], "exists": bool(fs),
                    "name": "・".join(f.name for f in fs) if fs else "(無いので同梱の VBA の絵を出します)"})
    return out
