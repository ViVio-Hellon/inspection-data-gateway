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
    """使用パスをオフライン・全部 `folder`(省くと基準データ)に向ける(仕掛台帳は第1・第2 とも同じフォルダ)。"""
    folder = folder or build()
    paths.set_mode(paths.OFFLINE)
    f = str(folder) + "/"
    user_settings.save(paths.KEY_PATHS, {paths.KEY_ITEMS: {k: f for k in ("quality", "packing", "history", "factory")},
                                         paths.KEY_LOT: {paths.FIRST: f, paths.SECOND: f}})
    return folder


def copy_to(tmp: Path) -> Path:
    """書き換える試験用に、基準データを `tmp` へ写す。"""
    src = build()
    tmp.mkdir(parents=True, exist_ok=True)
    for p in src.glob("*.sqlite3"):
        shutil.copy2(p, tmp / p.name)
    return tmp


QF_HEAD = ("写真確認No", "LotNo", "材質", "調質", "空冷時間(秒)", "記入者")
QF_ROWS = (("101", "H9022K1", "75S", "T651", "12", "検査員A"),
           ("", "N8061T0", "14S", "T6", "15", "検査員B"),
           ("102", "L8281G0", "75S", "T6", "11", "検査員A"),
           ("103", "L8281G0", "75S", "T6", "10", "検査員C"))


def write_qf_book(folder: Path, rows=QF_ROWS, head=QF_HEAD) -> Path:
    """`<品質基準>/75S・14S焼入れチェック/QF空冷時間チェックリスト.xlsx`(Sheet1: 4 行目が見出し、5 行目から A=写真確認No、B=LotNo)。"""
    import zipfile
    from xml.sax.saxutils import escape
    m = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def row(n: int, vals) -> str:
        cells = "".join(f'<c r="{chr(65 + i)}{n}" t="inlineStr"><is><t>{escape(str(v))}</t></is></c>'
                        for i, v in enumerate(vals) if v != "")
        return f'<row r="{n}">{cells}</row>'

    body = row(1, ["QF空冷時間チェックリスト"]) + row(4, head) + "".join(row(5 + i, v) for i, v in enumerate(rows))
    p = Path(folder) / "75S・14S焼入れチェック" / "QF空冷時間チェックリスト.xlsx"
    p.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("xl/workbook.xml", f'<workbook xmlns="{m}" xmlns:r="{r}"><sheets>'
                                      f'<sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr("xl/worksheets/sheet1.xml", f'<worksheet xmlns="{m}"><sheetData>{body}</sheetData></worksheet>')
    return p
