"""COM受信テスト(保守 → 測定機器ID)。Windows 以外では擬似端末を受信機の代わりにする。"""
import os
import sys
import time
import unittest

from tests import _isolation  # noqa: F401
from ufkensa.kensa import comport


@unittest.skipIf(sys.platform == "win32", "擬似端末は Windows に無い")
class ComReadTest(unittest.TestCase):
    def tearDown(self) -> None:
        comport.stop()

    def test_lines_split_by_cr(self) -> None:
        import pty
        master, slave = pty.openpty()
        self.addCleanup(os.close, master)
        name = os.ttyname(slave)
        comport.start(name)
        os.write(master, b"DT10001+0000025.012M\r#L395.8\r")
        for _ in range(30):
            st = comport.status()
            if len(st["lines"]) >= 2:
                break
            time.sleep(0.1)
        texts = [x["text"] for x in st["lines"]]
        self.assertEqual(texts, ["DT10001+0000025.012M", "#L395.8"])
        self.assertTrue(st["lines"][0]["hex"].startswith("44 54 31"))
        self.assertEqual(comport.status(st["last"])["lines"], [])     # since で新しい分だけ
        os.close(slave)
        self.assertFalse(comport.stop()["running"])

    def test_bad_settings(self) -> None:
        with self.assertRaises(comport.ComError):
            comport.start("")
        with self.assertRaises(comport.ComError):
            comport.start("/dev/null", parity="X")
        with self.assertRaises(comport.ComError):
            comport.start("/無いポート")


if __name__ == "__main__":
    unittest.main()
