"""複数引当によるラテラルボーの検査(JudgeLotAllocations FormRelat:5257-5333 / SaveLoad:560 / UF1:678)

1 ロットに引当が 2 件(Lot情報履歴の行が 2 行)あり、片方の納入先だけが ラテフリ直角 で ラテラルボー=要 のとき、
「1 件でも要なら要」でラテラルボー欄(L11・Lateral?)が出て、入力必須(印刷前チェック)・判定・保存・印刷・
読み戻しに入ることを、利用者の操作どおり Session で確かめる。

合成データ(一時コピーだけに入れる。tests/fixtures は触らない):
- SIKAHIKI.仕掛 に H7345B0 の 2 件目の引当(受注 59999999)、SIKAODR.仕掛 にその受注(納入先「ﾃｽﾄﾉｳﾆﾕｳｻｷ(ｶ」は架空)
- 品質基準.ラテフリ直角 にその納入先 × 厚 15 / 巾 1252 / 丈 2502 で ラテラルボー=要 の行
もとの引当(ｸﾉｻﾝｷﾞﾖｳ)は ラテフリ直角 に該当なし = 不要。
"""
from __future__ import annotations

import sqlite3
from pathlib import Path

from tests.test_kensa_session import LOT, SessionBase
from ufkensa.kensa import wiring
from ufkensa.kensa.store import HOZON, LOT_HISTORY

NOU = "ﾃｽﾄﾉｳﾆﾕｳｻｷ(ｶ"
ORDER = "59999999"
FIVE = "SUCDE"


def add_allocation(db_dir) -> None:
    con = sqlite3.connect(db_dir / "SIKAHIKI.sqlite3")
    cols = [r[1] for r in con.execute('PRAGMA table_info("仕掛")')]
    row = dict(zip(cols, con.execute('SELECT * FROM "仕掛" WHERE "ﾛｯﾄ番号"=?', (LOT,)).fetchone()))
    row.update({"引当番号": "69999999", "受注番号": ORDER, "受注_納入先ｺｰﾄﾞ": "99999999",
                "受注_納入先名称": NOU, "引当数量": "1"})
    con.execute('INSERT INTO "仕掛" (' + ",".join(f'"{c}"' for c in row) + ") VALUES ("
                + ",".join("?" * len(row)) + ")", list(row.values()))
    con.commit()
    con.close()
    # 受注(納入先は受注の 納入先名称 から Lot情報履歴 へ入る)
    con = sqlite3.connect(db_dir / "SIKAODR.sqlite3")
    cols = [r[1] for r in con.execute('PRAGMA table_info("仕掛")')]
    row = dict(zip(cols, con.execute('SELECT * FROM "仕掛" WHERE "受注番号"=?', ("54033469",)).fetchone()))
    row.update({"受注番号": ORDER, "納入先名称": NOU})
    con.execute('INSERT INTO "仕掛" (' + ",".join(f'"{c}"' for c in row) + ") VALUES ("
                + ",".join("?" * len(row)) + ")", list(row.values()))
    con.commit()
    con.close()
    con = sqlite3.connect(db_dir / "品質基準.sqlite3")
    cols = [r[1] for r in con.execute('PRAGMA table_info("ラテフリ直角")')]
    vals = {"管理番号": "9001", "納入先": NOU, "用途コード": "", "製造材質": "", "製造調質": "",
            "厚下": "10", "厚上": "20", "製造板厚": "10≦ｔ≦20", "幅下": "1000", "幅上": "1500",
            "製造板幅": "1000≦w≦1500", "丈下": "2000", "丈上": "3000", "製造板丈": "2000≦l≦3000",
            "ラテラルボー": "要", "公差": "5", "備考": "合成"}
    vals = {k: v for k, v in vals.items() if k in cols}
    con.execute('INSERT INTO "ラテフリ直角" (' + ",".join(f'"{c}"' for c in vals) + ") VALUES ("
                + ",".join("?" * len(vals)) + ")", list(vals.values()))
    con.commit()
    con.close()


class Base(SessionBase):
    multi = True

    def setUp(self) -> None:
        super().setUp()
        if self.multi:
            add_allocation(self.db_dir)

    def lateral_visible(self) -> dict[str, bool]:
        st = self.s.state
        return {n: st.ctl(n).visible for n in ["L11"] + ["Lateral" + p for p in FIVE]}

    def history_rows(self):
        return [r for r in wiring.sheet_rows(self.store, LOT_HISTORY)[1:] if r[2] == LOT]


class SingleAllocationTest(Base):
    """対照: 合成の引当なし → ラテラル不要。"""
    multi = False

    def test_not_required(self) -> None:
        self.register()
        self.assertEqual(len(self.history_rows()), 1)
        self.assertEqual(self.lateral_visible(), {n: False for n in self.lateral_visible()})
        self.assertFalse(self.lateral_visible()["L11"])


class MultiAllocationLateralTest(Base):
    def test_allocations_and_visibility(self) -> None:
        self.register()
        rows = self.history_rows()
        self.assertEqual(sorted(r[12] for r in rows), sorted([NOU, "ｸﾉｻﾝｷﾞﾖｳ(ｶ ﾎﾝｼﾔ"]))
        st = self.s.state
        vis = self.lateral_visible()
        self.assertTrue(vis["L11"], (vis, st.additional, st.lateral_text))
        self.assertTrue(vis["LateralS"] and vis["LateralC"] and vis["LateralE"])
        self.assertEqual(st.lateral_text, "5.0")
        self.assertNotIn("ﾗﾃﾗﾙﾎﾞｰ検査不要", st.additional)

    def test_flow_print_save_reload(self) -> None:
        self.register()
        s = self.s
        self.fill("DT")
        self.fill("#L")
        self.strain_visual_height()
        self.ok(s.min_max())
        # ラテラル未入力 → 印刷前チェックで止まる(MsgError(ラテ & ":ﾗﾃﾗﾙﾎﾞｰ入力なし") UF1:3688)
        r = s.print_check()
        self.assertIn("S:ﾗﾃﾗﾙﾎﾞｰ入力なし", [m.text for m in r.messages])
        self.assertFalse(s.print_ready)
        lat = [p for p in FIVE if s.state.ctl("Lateral" + p).visible]
        self.assertEqual(lat, ["S", "C", "E"])            # 検番 H は U/D を出さない(LSelectLateral FormRelat:5335-)
        # 判定(LateralbowJudgColor。公差 5.0 はラテラルが要の引当の行から)
        r = self.ok(s.set_value("LateralS", "9.0"))
        self.assertEqual(s.state.ctl("LateralS").color, "NG")
        self.assertIn("検査ﾃﾞｰﾀに公差外有り！", [m.text for m in r.messages])
        self.ok(s.set_value("UFConfi.TB1", "許容"))                     # 確認(公差外)に答えてから次へ
        for p in lat:
            self.ok(s.set_value("Lateral" + p, "1.0"))
            self.assertEqual(s.state.ctl("Lateral" + p).color, "OK")
        # 作業中保存からの読み戻し(Load_Text → LSelectLateral)
        s2 = self.new_session()
        self.ok(s2.start())
        self.s = s = s2
        self.assertTrue(self.lateral_visible()["L11"])
        self.assertEqual(s.state.lateral_text, "5.0")
        self.assertEqual([s.state.ctl("Lateral" + p).value for p in lat], ["1.0"] * 3)
        # 印刷・保存(MainDataSave の 72 列〜 = Sﾗﾃﾗﾙﾎﾞｰ…)
        self.ok(s.min_max())
        self.ok(s.print_check())
        self.assertTrue(s.print_ready)
        self.ok(s.print_commit({"print": True}))
        hz = wiring.sheet_rows(self.store, HOZON)
        saved = {h: v for h, v in zip(hz[0], hz[-1]) if "ﾗﾃﾗﾙﾎﾞｰ" in str(h)}
        self.assertEqual(hz[-1][1], LOT)
        self.assertEqual([saved[p + "ﾗﾃﾗﾙﾎﾞｰ"] for p in lat], [1.0] * 3)
        html = Path(s.report["report"]["html"]).read_text(encoding="utf-8")
        self.assertIn("ﾗﾃﾗﾙ", html)

    def test_order_independent(self) -> None:
        """引当の並び(要の行が先/後)によらず要。公差・Sumisei は要の行の照合結果が残る(FormRelat:5298-5325)。"""
        from ufkensa.kensa import layout
        self.register()
        s = self.s
        rows = s.hooks.get_lot_info(s.state, LOT)
        self.assertEqual(len(rows), 2)
        for order in (rows, rows[::-1]):
            s.state.lateral_text = ""
            r = layout.judge_lot_allocations(s.state, [list(x) for x in order], s.hooks)
            self.assertTrue(r.need_lateral)
            self.assertEqual(r.ken_no, "H")
            self.assertEqual(s.state.lateral_text, "5.0")
