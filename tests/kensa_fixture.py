"""試験用の DB 置き場(受領した Access から作った基準データ)

`tests/fixtures/kensa/<DB名>/` の TSV(bridge.vbs と同じ形。`scripts/dev_accdb_to_tsv.py` で作った)を
一時フォルダに .sqlite3 にして置き、使用パス(オフライン)をそこへ向ける。

- 品質基準・伝送用ファイル: 受領ファイルの全表(~TMP 表を除く)。名簿は氏名・カード番号を伏せた
- SIKALOTNOW / SIKAHIKINOW / SIKAODRNOW: 108 ロット分だけ抜き出したもの
"""
from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from typing import Optional

from ufkensa import user_settings
from ufkensa.accdb_bridge import importer
from ufkensa.kensa import paths

FIXTURES = Path(__file__).resolve().parent / "fixtures" / "kensa"
NAMES = ("品質基準", "伝送用ファイル", "SIKALOTNOW", "SIKAHIKINOW", "SIKAODRNOW")

_built: Optional[Path] = None


def file_name(name: str) -> str:
    """基準データの DB 名 → ファイル名(仕掛は SIKALOT.sqlite3 など。paths.DB_FILES と同じ)。"""
    f = paths.DB_FILES.get(name, ("", ""))[1]
    return f if f and "\\" not in f else f"{name}.sqlite3"


def build() -> Path:
    """基準データの .sqlite3 を作る(1 回だけ)。置いたフォルダを返す。"""
    global _built
    if _built and _built.exists():
        return _built
    out = Path(tempfile.mkdtemp(prefix="ufkensa_kensa_db_"))
    for name in NAMES:
        importer.import_tsv_dir(FIXTURES / name, out / file_name(name))
    _built = out
    return out


def use_offline(folder: Optional[Path] = None) -> Path:
    """使用パスをオフライン・全部 `folder`(省くと基準データ)に向ける(第1・第2 とも同じフォルダ)。"""
    folder = folder or build()
    paths.set_mode(paths.OFFLINE)
    one = {k: str(folder) + "/" for k in (paths.LOT, "quality", "packing", "history", "factory")}
    user_settings.save(paths.KEY_PATHS, {paths.FIRST: dict(one), paths.SECOND: dict(one)})
    return folder


def copy_to(tmp: Path) -> Path:
    """書き換える試験用に、基準データを `tmp` へ写す。"""
    src = build()
    tmp.mkdir(parents=True, exist_ok=True)
    for p in src.glob("*.sqlite3"):
        shutil.copy2(p, tmp / p.name)
    return tmp
