# -*- coding: utf-8 -*-
"""
تشغيل نظام الأفراد على أي حاسبة: البحث عن الملفات والمسارات
بدون الاعتماد على قرص H: أو مسار ثابت.
"""

from __future__ import annotations

import configparser
import os
import shutil
import stat
import string
import subprocess
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

ROOT = Path(__file__).resolve().parent
CFG_PATH = ROOT / "afrad_local.cfg"

MFILE_NAMES = (
    "MFILE.DBF",
    "MFILE.dbf",
    "mfile.dbf",
    "MFILE_updated.DBF",
    "MFILE_updated.dbf",
)
EXCEL_NAMES = (
    "New Microsoft Excel Worksheet.xlsx",
    "New Microsoft Excel Worksheet.xls",
)
SS_NAMES = ("ss.xls", "ss.xlsx", "SS.xls", "SS.xlsx")
MM_NAMES = ("mm.xlsx", "mm.xls", "MM.xlsx", "MM.xls")
FF_NAMES = ("ff.xlsx", "ff.xls", "FF.xlsx", "FF.xls")

EXE_SKIP = {
    "python.exe",
    "pythonw.exe",
    "py.exe",
    "pip.exe",
    "cmd.exe",
    "powershell.exe",
    "pwsh.exe",
    "winget.exe",
    "installer.exe",
    "setup.exe",
    "install.exe",
    "uninstall.exe",
    "git.exe",
    "node.exe",
    "code.exe",
}
EXE_HINTS = (
    "afrad",
    "afrd",
    "nizam",
    "nitham",
    "personnel",
    "employee",
    "emp",
    "main",
    "mfile",
    "fox",
)

VFP_RUNTIME_FILES = (
    "VFP9R.DLL",
    "VFP9T.DLL",
    "VFP9RENU.DLL",
    "VFP9RARA.DLL",
    "VFP8R.DLL",
    "VFP8RENU.DLL",
    "VFP7R.DLL",
    "VFP7RENU.DLL",
    "VFP6R.DLL",
    "VFP6RENU.DLL",
    "VFP500.DLL",
    "VFP5.DLL",
    "VFP5ENU.DLL",
    "MSVCR71.DLL",
    "MSVCR70.DLL",
    "gdiplus.dll",
)


def setup_stdio() -> None:
    """UTF-8 على ويندوز حتى لا تظهر رموز خطأ بدل العربي."""
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    if os.name == "nt":
        try:
            subprocess.run(["chcp", "65001"], capture_output=True, check=False)
        except Exception:
            pass


def project_root() -> Path:
    return ROOT


def _unique_existing(paths: Iterable[Path]) -> List[Path]:
    seen = set()
    out: List[Path] = []
    for raw in paths:
        if raw is None:
            continue
        try:
            path = Path(os.path.expandvars(str(raw))).expanduser()
        except Exception:
            continue
        try:
            if not path.exists():
                continue
            key = str(path.resolve()).lower()
        except Exception:
            key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(path)
    return out


def iter_search_roots() -> List[Path]:
    home = Path.home()
    candidates: List[Path] = []
    env_home = os.environ.get("AFRAD_HOME")
    if env_home:
        candidates.append(Path(env_home))
    candidates.extend(
        [
            ROOT,
            ROOT / "data",
            home / "Desktop",
            home / "Desktop" / "Afrad2_work",
            home / "Documents",
            home / "Downloads",
            home / "OneDrive" / "Desktop",
            Path(r"C:\Afrad2_work"),
            Path(r"C:\Afrad"),
            Path(r"C:\AFRAD"),
            Path(r"D:\Afrad2_work"),
            Path(r"D:\Afrad"),
            Path("/mnt/h"),
            Path("/mnt/c/Afrad2_work"),
        ]
    )
    if os.name == "nt":
        for letter in string.ascii_uppercase:
            candidates.append(Path(f"{letter}:\\"))
    return _unique_existing(candidates)


def load_cfg() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    if CFG_PATH.exists():
        try:
            cfg.read(CFG_PATH, encoding="utf-8")
        except Exception:
            cfg.read(CFG_PATH, encoding="cp1256")
    if "paths" not in cfg:
        cfg["paths"] = {}
    return cfg


def save_cfg_value(key: str, value: Path) -> None:
    try:
        cfg = load_cfg()
        cfg["paths"][key] = str(value)
        with CFG_PATH.open("w", encoding="utf-8") as fh:
            cfg.write(fh)
    except Exception:
        pass


def cfg_path(key: str) -> Optional[Path]:
    cfg = load_cfg()
    raw = cfg["paths"].get(key, "").strip()
    if not raw:
        return None
    path = Path(raw)
    return path if path.exists() else None


def find_file(names: Sequence[str], key: Optional[str] = None) -> Optional[Path]:
    if key:
        cached = cfg_path(key)
        if cached:
            return cached
    wanted = [n.lower() for n in names]
    for root in iter_search_roots():
        try:
            entries = list(root.iterdir())
        except Exception:
            continue
        files = {p.name.lower(): p for p in entries if p.is_file()}
        for name in wanted:
            hit = files.get(name)
            if hit:
                if key:
                    save_cfg_value(key, hit)
                return hit
    return None


def find_mfile_dbf() -> Optional[Path]:
    return find_file(MFILE_NAMES, key="dbf")


def find_excel() -> Optional[Path]:
    found = find_file(EXCEL_NAMES, key="excel")
    if found:
        return found
    # أي ملف إكسل واضح في مجلد المشروع أو data
    for root in (ROOT, ROOT / "data"):
        if not root.exists():
            continue
        for pat in ("*.xlsx", "*.xls"):
            for hit in sorted(root.glob(pat)):
                if hit.name.startswith("~$"):
                    continue
                save_cfg_value("excel", hit)
                return hit
    return None


def find_named(names: Sequence[str], key: str) -> Optional[Path]:
    return find_file(names, key=key)


def companion_files(src: Path) -> List[Path]:
    files = [src]
    stem = src.with_suffix("")
    for ext in (".CDX", ".cdx", ".FPT", ".fpt", ".DBT", ".dbt", ".IDX", ".idx"):
        side = Path(str(stem) + ext)
        if side.exists():
            files.append(side)
    return files


def ensure_local_mfile() -> Optional[Path]:
    """نسخة عمل محلية من MFILE حتى لا يعتمد البرنامج على H:."""
    local = ROOT / "MFILE.DBF"
    if local.exists():
        return local
    data_local = ROOT / "data" / "MFILE.DBF"
    if data_local.exists():
        try:
            shutil.copy2(data_local, local)
            for side in companion_files(data_local)[1:]:
                shutil.copy2(side, ROOT / side.name)
        except Exception:
            return data_local
        return local
    found = find_mfile_dbf()
    if not found:
        return None
    if found.resolve() == local.resolve():
        return local
    try:
        for item in companion_files(found):
            dest = ROOT / item.name
            if item.suffix.upper() == ".DBF":
                dest = local
            if not dest.exists():
                shutil.copy2(item, dest)
        return local if local.exists() else found
    except Exception:
        return found


def make_data_writable(root: Path) -> None:
    if not root.exists():
        return
    for pat in ("*.DBF", "*.dbf", "*.CDX", "*.cdx", "*.FPT", "*.fpt", "*.IDX", "*.idx"):
        for path in root.glob(pat):
            try:
                os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
            except Exception:
                pass


def ensure_drive_letter(letter: str, target: Path) -> str:
    """
    ربط حرف قرص قديم (غالباً H:) بالمجلد الحالي.
    أنظمة الفوكس القديمة تبحث عن H:\\MFILE.DBF حتى لو نُسخ المجلد لجهاز آخر.
    """
    if os.name != "nt":
        return "not-windows"
    drive = f"{letter}:"
    drive_root = Path(f"{drive}\\")
    if drive_root.exists():
        return "exists"
    try:
        result = subprocess.run(
            ["subst", drive, str(target)],
            capture_output=True,
            text=True,
            check=False,
        )
        if drive_root.exists():
            return "mapped"
        if result.returncode != 0:
            return "failed"
    except Exception:
        return "failed"
    return "failed"


def write_config_fpw(app_dir: Path) -> Path:
    temp = app_dir / "TEMP"
    try:
        temp.mkdir(exist_ok=True)
    except Exception:
        pass
    content = (
        "DEFAULT=.\n"
        "PATH=.;.\\data\n"
        "RESOURCE=OFF\n"
        "SCREEN=ON\n"
        "TMPFILES=.\\TEMP\n"
        "SORTWORK=.\\TEMP\n"
        "EDITWORK=.\\TEMP\n"
        "PROGWORK=.\\TEMP\n"
        "MVCOUNT=8192\n"
        "COMMAND=SET EXCLUSIVE OFF\n"
    )
    cfg = app_dir / "config.fpw"
    bak = app_dir / "config.fpw.bak_portable"
    try:
        if cfg.exists() and not bak.exists():
            shutil.copy2(cfg, bak)
        cfg.write_text(content, encoding="ascii", errors="replace")
    except Exception:
        pass
    return cfg


def _score_exe(path: Path) -> int:
    name = path.name.lower()
    if name in EXE_SKIP:
        return -1
    score = 0
    for hint in EXE_HINTS:
        if hint in name:
            score += 10
    if name.endswith(".exe"):
        score += 1
    if path.parent.resolve() == ROOT:
        score += 5
    return score


def find_personnel_exe() -> Optional[Path]:
    cached = cfg_path("personnel_exe")
    if cached and cached.suffix.lower() == ".exe" and cached.name.lower() not in EXE_SKIP:
        return cached
    found: List[Tuple[int, Path]] = []
    search_dirs = [ROOT, ROOT / "bin", ROOT / "app"]
    search_dirs.extend(iter_search_roots()[:12])
    seen = set()
    for folder in search_dirs:
        try:
            resolved = folder.resolve()
        except Exception:
            resolved = folder
        key = str(resolved).lower()
        if key in seen:
            continue
        seen.add(key)
        try:
            children = list(folder.iterdir())
        except Exception:
            continue
        for item in children:
            if item.is_file() and item.suffix.lower() == ".exe":
                score = _score_exe(item)
                if score >= 0:
                    found.append((score, item))
            elif item.is_dir() and item.name.lower() not in {
                ".git",
                "__pycache__",
                "tests",
                "node_modules",
            }:
                try:
                    for nested in item.glob("*.exe"):
                        score = _score_exe(nested)
                        if score >= 0:
                            found.append((score, nested))
                except Exception:
                    continue
    if not found:
        return None
    found.sort(key=lambda x: x[0], reverse=True)
    hinted = [item for item in found if item[0] >= 10]
    if hinted:
        best = hinted[0][1]
        save_cfg_value("personnel_exe", best)
        return best
    try:
        root_res = ROOT.resolve()
    except Exception:
        root_res = ROOT
    local = [
        item
        for item in found
        if item[1].parent.resolve() == root_res and item[1].name.lower() not in EXE_SKIP
    ]
    if len(local) == 1:
        best = local[0][1]
        save_cfg_value("personnel_exe", best)
        return best
    return None


def copy_vfp_runtime(app_dir: Path) -> List[Path]:
    copied: List[Path] = []
    search = [
        Path(r"C:\Program Files (x86)\Common Files\Microsoft Shared\VFP"),
        Path(r"C:\Program Files\Common Files\Microsoft Shared\VFP"),
        Path(os.environ.get("WINDIR", r"C:\Windows")) / "SysWOW64",
        Path(os.environ.get("WINDIR", r"C:\Windows")) / "System32",
        ROOT,
    ]
    for name in VFP_RUNTIME_FILES:
        dest = app_dir / name
        if dest.exists():
            continue
        for folder in search:
            src = folder / name
            if not src.exists():
                alt = folder / name.lower()
                src = alt if alt.exists() else src
            if src.exists() and src.resolve() != dest.resolve():
                try:
                    shutil.copy2(src, dest)
                    copied.append(dest)
                    break
                except Exception:
                    continue
    return copied


def launch_personnel(exe: Optional[Path] = None) -> int:
    exe = exe or find_personnel_exe()
    if exe is None or not exe.exists():
        print("لم يتم العثور على برنامج نظام الأفراد (ملف EXE).")
        print("انسخ مجلد البرنامج كاملاً من الحاسبة التي يعمل عليها إلى هذا المجلد:")
        print(f"  {ROOT}")
        print("ثم انقر نقراً مزدوجاً على START.bat")
        print("الملفات المهمة: البرنامج EXE + VFP*.DLL + ملفات DBF/CDX")
        return 0
    app_dir = exe.parent
    make_data_writable(app_dir)
    make_data_writable(ROOT)
    make_data_writable(ROOT / "data")
    ensure_local_mfile()
    mapped = ensure_drive_letter("H", ROOT)
    write_config_fpw(app_dir)
    copy_vfp_runtime(app_dir)
    env = os.environ.copy()
    env["PATH"] = str(app_dir) + os.pathsep + env.get("PATH", "")
    cfg = app_dir / "config.fpw"
    args = [str(exe)]
    if cfg.exists():
        args.extend(["-c", str(cfg)])
    print(f"تشغيل نظام الأفراد: {exe}")
    print(f"المجلد: {app_dir}")
    if mapped == "mapped":
        print("تم ربط القرص H: بهذا المجلد حتى تعمل المسارات القديمة.")
    try:
        subprocess.Popen(args, cwd=str(app_dir), env=env)
        return 0
    except OSError as exc:
        print("تعذر فتح البرنامج:", exc)
        print("انسخ مكتبات FoxPro (VFP9R.DLL أو VFP6R.DLL) من الحاسبة التي يعمل عليها البرنامج")
        print("وضعها في نفس مجلد EXE ثم أعد تشغيل START.bat")
        return 0


def prepare_environment() -> Dict[str, str]:
    """تحضير الجهاز الحالي بدون إظهار أخطاء للمستخدم."""
    status = {
        "root": str(ROOT),
        "mfile": "",
        "excel": "",
        "exe": "",
        "hdrive": ensure_drive_letter("H", ROOT),
    }
    try:
        make_data_writable(ROOT)
        make_data_writable(ROOT / "data")
        mfile = ensure_local_mfile()
        if mfile:
            status["mfile"] = str(mfile)
        excel = find_excel()
        if excel:
            status["excel"] = str(excel)
        exe = find_personnel_exe()
        if exe:
            status["exe"] = str(exe)
            write_config_fpw(exe.parent)
            copy_vfp_runtime(exe.parent)
        else:
            write_config_fpw(ROOT)
    except Exception as exc:
        status["note"] = str(exc)
    return status


def default_update_paths() -> Tuple[Path, Path]:
    excel = find_excel() or (ROOT / "data" / "New Microsoft Excel Worksheet.xlsx")
    dbfp = ensure_local_mfile() or find_mfile_dbf() or (ROOT / "MFILE.DBF")
    return excel, dbfp
