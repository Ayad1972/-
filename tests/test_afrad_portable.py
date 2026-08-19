# -*- coding: utf-8 -*-
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import afrad_portable as ap
import start_afrad
import update_mfile_from_excel as upd
BAT_FILES = [
    "START.bat",
    "OPEN.bat",
    "GO.bat",
    "_find_python.bat",
    "install_python.bat",
    "run_update_mfile.bat",
    "run_merge_mm_into_ss.bat",
    "run_merge_ff_into_ss.bat",
    "copy_to_desktop.bat",
    "download_mfile_to_pc.bat",
    "speed_up_pc.bat",
]


class AsciiBatTests(unittest.TestCase):
    def test_launchers_are_ascii_without_bom(self):
        for name in BAT_FILES:
            path = ROOT / name
            data = path.read_bytes()
            self.assertFalse(data.startswith(b"\xef\xbb\xbf"), name)
            data.decode("ascii")


class PathFinderTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="afrad_test_"))
        self.addCleanup(self._cleanup)
        src = ROOT / "MFILE_updated.DBF"
        (self.tmp / "MFILE.DBF").write_bytes(src.read_bytes())
        (self.tmp / "New Microsoft Excel Worksheet.xlsx").write_bytes(b"PK\x03\x04fake")
        self._old_root = ap.ROOT
        self._old_cfg = ap.CFG_PATH
        ap.ROOT = self.tmp
        ap.CFG_PATH = self.tmp / "afrad_local.cfg"
        self._old_roots = ap.iter_search_roots
        ap.iter_search_roots = lambda: [self.tmp]

    def _cleanup(self):
        ap.ROOT = self._old_root
        ap.CFG_PATH = self._old_cfg
        ap.iter_search_roots = self._old_roots
        for item in sorted(self.tmp.rglob("*"), reverse=True):
            try:
                if item.is_file():
                    item.unlink()
                else:
                    item.rmdir()
            except OSError:
                pass
        try:
            self.tmp.rmdir()
        except OSError:
            pass

    def test_finds_local_mfile_without_h_drive(self):
        found = ap.find_file(ap.MFILE_NAMES, key="dbf")
        self.assertIsNotNone(found)
        self.assertEqual(found.name.upper(), "MFILE.DBF")
        self.assertTrue(str(found).startswith(str(self.tmp)))

    def test_ensure_local_mfile_uses_project_copy(self):
        path = ap.ensure_local_mfile()
        self.assertEqual(path, self.tmp / "MFILE.DBF")

    def test_default_update_paths_not_h_drive(self):
        excel, dbfp = ap.default_update_paths()
        self.assertNotEqual(str(excel)[:2].upper(), "H:")
        self.assertNotEqual(str(dbfp)[:2].upper(), "H:")

    def test_score_prefers_afrad_exe(self):
        afrad = self.tmp / "Afrad.exe"
        other = self.tmp / "setup.exe"
        afrad.write_bytes(b"MZ")
        other.write_bytes(b"MZ")
        self.assertGreater(ap._score_exe(afrad), ap._score_exe(other))

    def test_write_config_fpw_ascii(self):
        cfg = ap.write_config_fpw(self.tmp)
        text = cfg.read_text(encoding="ascii")
        self.assertIn("RESOURCE=OFF", text)
        self.assertIn("DEFAULT=.", text)

    def test_does_not_pick_unrelated_exe(self):
        (self.tmp / "setup.exe").write_bytes(b"MZ")
        self.assertIsNone(ap.find_personnel_exe())
        (self.tmp / "Afrad2.exe").write_bytes(b"MZ")
        found = ap.find_personnel_exe()
        self.assertIsNotNone(found)
        self.assertEqual(found.name, "Afrad2.exe")


class ScriptBehaviorTests(unittest.TestCase):
    def test_inspect_missing_excel_does_not_crash(self):
        dbf = ROOT / "MFILE_updated.DBF"
        code = upd.inspect_files(Path("/tmp/does-not-exist.xlsx"), dbf)
        self.assertEqual(code, 0)

    def test_inspect_missing_dbf_returns_1_not_traceback(self):
        code = upd.inspect_files(Path("/tmp/does-not-exist.xlsx"), Path("/tmp/no.dbf"))
        self.assertEqual(code, 1)

    def test_update_main_missing_files_arabic_exit_1(self):
        code = upd.main(
            [
                "--excel",
                "/tmp/missing.xlsx",
                "--dbf",
                "/tmp/missing.dbf",
                "--dry-run",
            ]
        )
        self.assertEqual(code, 1)

    def test_start_check_always_zero(self):
        code = start_afrad.main(["--check"])
        self.assertEqual(code, 0)

    def test_default_paths_find_repo_mfile(self):
        excel, dbfp = upd.default_paths()
        self.assertTrue(dbfp.exists(), dbfp)
        self.assertNotEqual(str(dbfp).replace("\\", "/"), "H:/MFILE.DBF")


if __name__ == "__main__":
    unittest.main(verbosity=2)
