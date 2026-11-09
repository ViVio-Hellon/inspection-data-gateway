"""判定の画像(VBA の UFOK / UFNG / UF対角 フォームの絵。2026-11-09)

VBA は絵をフォームに埋め込み、判定のたびに出していた:
- UFOK / UFNG(UFOK.frm:15-26 / UFNG.frm:16-25): 合格・不合格で `Show vbModeless`、3 秒で閉じる
  (全ヵ所検査 ZenSh2.cls:78-100・Cask.bas:1429-1450。音楽合格 / 音楽不合格 と同時)
- UF対角(UF対角.frm:12-50): 対角線の欄に値が入ったとき(Cask.bas:1899 UF対角起動。音楽対角 と同時)、
  DB1〜DB5 のうち 1 枚(検査シートは行 14〜18 で決まり、ほかは乱数)を 2 秒出す

ここでは、使用パス「合格の画像」「不合格の画像」「対角線の画像1〜5」(2026-11-09: 絵ごとのファイル。
前は「判定の画像」フォルダに決まった名前で置いた。そのフォルダは 1 回だけ枠へ読み替える)の画像を配る。
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


#: 2026-11-09: フォルダをやめ、絵ごとにファイルを決める(設定 → 使用パス)。key → (種類, 同梱の名前, 表示名)
SLOTS = {
    "judge_ok": ("ok", "OK.jpg", "合格の画像"),
    "judge_ng": ("ng", "NG.jpg", "不合格の画像"),
    **{f"judge_diag{i}": ("diag", f"対角{i}.jpg", f"対角線の画像{i}") for i in range(1, 6)},
}
#: 前の設定(フォルダ)。読み替えのためだけに残す
OLD_KEY = "judge_images"
MISSING_NOTE = "ファイルがありません(同梱の VBA の絵を出します)"


def bundled(slot: str) -> Path:
    return DEFAULT_DIR / SLOTS[slot][1]


def configured(slot: str) -> str:
    from . import paths
    return paths.get(slot) if slot in SLOTS else ""


def _usable(p: Path) -> bool:
    try:
        return p.is_file() and p.suffix.lower() in EXT
    except OSError:
        return False


def resolve(slot: str) -> tuple[Path, str]:
    """その枠で出すファイルと、その出どころ("set" / "missing" / "default")。"""
    from . import paths
    text = configured(slot)
    if not text:
        return bundled(slot), "default"
    try:
        p = Path(paths.to_local(text))
    except Exception:  # noqa: BLE001
        return bundled(slot), "missing"
    return (p, "set") if _usable(p) else (bundled(slot), "missing")


def slots_of(kind: str) -> list[str]:
    """出す枠。対角線は決めた枠だけ(1 つも決めていなければ同梱の 対角1〜5。VBA の DB1〜DB5 を乱数で)。"""
    keys = [k for k, v in SLOTS.items() if v[0] == kind]
    if kind == "diag":
        chosen = [k for k in keys if configured(k)]
        return chosen or keys
    return keys


def slot_info(slot: str) -> dict:
    """設定画面の見本用。"""
    f, src = resolve(slot)
    return {"slot": slot, "label": SLOTS[slot][2], "path": configured(slot), "file": f.name, "source": src,
            "note": MISSING_NOTE if src == "missing" else ("既定(同梱の VBA の絵)" if src == "default" else "")}


def migrate_from_folder(folder_text: str) -> dict[str, str]:
    """前の「判定の画像」フォルダ → 枠ごとのファイル(決まった名前、無ければ名前に含む言葉で)。"""
    from . import paths
    try:
        folder = Path(paths.to_local(folder_text))
    except Exception:  # noqa: BLE001
        return {}
    out: dict[str, str] = {}
    for kind in KINDS:
        fs = files(kind, folder)
        keys = [k for k, v in SLOTS.items() if v[0] == kind]
        for k, f in zip(keys, fs):
            out[k] = paths.join(folder_text, f.name)
    return out


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
    if not out:
        out = _loose(folder, kind)
    return out


#: 2026-11-09: 決まった名前でなくても、名前に含む言葉で見分ける(合格・不合格・対角)。NG・不合格 を先に見る
WORDS = {"ng": ("NG", "不合格", "ﾌｺﾞｳｶｸ", "ダメ", "NO"), "ok": ("OK", "合格", "ｺﾞｳｶｸ"), "diag": ("対角", "ﾀｲｶｸ", "DIAG")}


def _loose(folder: Path, kind: str) -> list[Path]:
    try:
        fs = sorted(f for f in folder.iterdir() if f.is_file() and f.suffix.lower() in EXT)
    except OSError:
        return []

    def kind_of(f: Path) -> str:
        stem = f.stem.upper()
        for k in ("ng", "diag", "ok"):
            if any(w.upper() in stem for w in WORDS[k]):
                return k
        return ""
    return [f for f in fs if kind_of(f) == kind][: (5 if kind == "diag" else 1)]


def effective(kind: str) -> list[Path]:
    """出す画像(枠ごとに決めたファイル。空・無いときは同梱の絵)。"""
    return [resolve(k)[0] for k in slots_of(kind)]


def available(folder: Optional[Path] = None) -> dict[str, list[str]]:
    """{種類: [枠…]}。2026-11-09: ファイル名ではなく枠(judge_ok など)を返す。folder を渡すと前の形(名前)。"""
    if folder is None:
        return {k: slots_of(k) for k in KINDS}
    return {k: [f.name for f in files(k, folder)] for k in KINDS}


def find(kind: str, name: str) -> Optional[Path]:
    """配ってよいファイル。name は枠(judge_ok など)。その種類の枠だけ。"""
    if name in SLOTS and SLOTS[name][0] == kind and name in slots_of(kind):
        return resolve(name)[0]
    return None


def listing(folder: Path) -> list[dict]:
    """設定画面の確認用。"""
    out = []
    for k in KINDS:
        fs = files(k, folder)
        out.append({"kind": k, "label": LABELS[k], "exists": bool(fs),
                    "name": "・".join(f.name for f in fs) if fs else "(無いので同梱の VBA の絵を出します)"})
    return out


def unused(folder: Path) -> list[str]:
    """フォルダにある画像のうち、どの種類にも使われないもの(設定の確認で知らせる)。"""
    used = {f.name for k in KINDS for f in files(k, folder)}
    try:
        return sorted(f.name for f in folder.iterdir() if f.is_file() and f.suffix.lower() in EXT and f.name not in used)
    except OSError:
        return []
