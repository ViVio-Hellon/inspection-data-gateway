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

    def test_ddm100_frames_and_ack(self) -> None:
        """DDM-100: STX〜ETX で 1 件。パケット形式は #L と数値、数値のみ形式も #L として渡し、ACK を返す。"""
        import pty
        import tty
        master, slave = pty.openpty()
        tty.setraw(slave)                        # 実際の COM と同じく、0x03 を Ctrl+C として扱わない
        self.addCleanup(os.close, master)
        got = []
        comport.start(os.ttyname(slave), sink=got.append, slot="ddm")
        self.addCleanup(comport.stop, "ddm")
        os.write(master, b"\x02#L\x05125.3\x03\x02395.8\x03")
        for _ in range(30):
            if len(got) >= 2:
                break
            time.sleep(0.1)
        self.assertEqual(got, ["#L125.3", "#L395.8"])
        time.sleep(0.2)
        self.assertEqual(os.read(master, 16), b"\x06\x06")          # 1 件ごとに ACK
        os.close(slave)

    def test_two_receivers_at_once(self) -> None:
        import pty
        m1, s1 = pty.openpty()
        m2, s2 = pty.openpty()
        import tty
        tty.setraw(s2)
        for fd in (m1, m2):
            self.addCleanup(os.close, fd)
        a, b = [], []
        comport.start(os.ttyname(s1), sink=a.append, slot="kensa:1")
        comport.start(os.ttyname(s2), sink=b.append, slot="kensa:2")
        self.addCleanup(comport.stop, None)
        os.write(m1, b"DT10001+00000025.01M\r")
        os.write(m2, b"\x02#L\x05395.8\x03")
        for _ in range(30):
            if a and b:
                break
            time.sleep(0.1)
        self.assertEqual((a, b), (["DT10001+00000025.01M"], ["#L395.8"]))
        self.assertEqual(len(comport.receivers()), 2)
        os.close(s1)
        os.close(s2)

    def test_bad_settings(self) -> None:
        with self.assertRaises(comport.ComError):
            comport.start("")
        with self.assertRaises(comport.ComError):
            comport.start("/dev/null", parity="X")
        with self.assertRaises(comport.ComError):
            comport.start("/無いポート")


if __name__ == "__main__":
    unittest.main()
