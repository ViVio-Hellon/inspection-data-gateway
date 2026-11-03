"""特殊検査モード(キャスク・全ヵ所〔一般 / 5N / 三菱受託〕・H494)

- `sheet`    … 作業シートとブックの模型(Excel のイベントの決まり)、Cask モジュールの共用の判定
- `cask`     … キャスク検査(UFCask・KySh1・梱包板番号 DB)
- `zenkasho` … 全ヵ所検査(UFMitsubishi P1・ZenSh1・Mitsubishi モジュール)
- `h494`     … H494 検査(UFMitsubishi P_H494・H494_SH1)

import すると 3 つの作業シートの Change / SelectionChange が `sheet.HANDLERS` に登録される。
"""
from . import sheet  # noqa: F401
from . import cask  # noqa: F401
from . import zenkasho  # noqa: F401
from . import h494  # noqa: F401
from .sheet import Book  # noqa: F401
