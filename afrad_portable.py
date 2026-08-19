# -*- coding: utf-8 -*-
"""
تشغيل نظام الأفراد على أي حاسبة: البحث عن الملفات والمسارات
بدون الاعتماد على قرص H: أو مسار ثابت، وبدون فحص أقراص فارغة.
"""

from __future__ import annotations

import configparser
import os
import shutil
import stat
import subprocess
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

ROOT = Path(__file__).resolve().parent
CFG_PATH = ROOT / "afrad_local.cfg"
LEGACY_DIR = Path(r"C:\Afrad2_work")

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
    """UTF-8 للطباعة إن أمكن. لا نغيّر صفحة أوامر ويندوز لأنها تسبب رموزاً تالفة."""
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def silence_windows_drive_errors() -> None:
    """منع نافذة 'أدخل قرصاً في المحرك' عند فحص مسار غير جاهز."""
    if os.name != "nt":
        return
    try:
        import ctypes

        ctypes.windll.kernel32.SetErrorMode(1)  # SEM_FAILCRITICALERRORS
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


def ready_windows_drives() -> List[Path]:
    """أقراص جاهزة فقط. لا نفحص A:/B: ولا محركات الأقراص الضوئية الفارغة."""
    if os.name != "nt":
        return []
    silence_windows_drive_errors()
    found: List[Path] = []
    try:
        import ctypes

        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        get_type = ctypes.windll.kernel32.GetDriveTypeW
        for i in range(26):
            if not bitmask & (1 << i):
                continue
            letter = chr(65 + i)
            if letter in "AB":
                continue
            root = f"{letter}:\\"
            dtype = int(get_type(root))
            # 2=removable 3=fixed 4=remote 6=ramdisk. Skip 5=CDROM unless already listed.
            if dtype in (2, 3, 4, 6):
                found.append(Path(root))
    except Exception:
        for letter in ("C", "D", "H"):
            p = Path(f"{letter}:\\")
            try:
                if p.exists():
                    found.append(p)
            except Exception:
                continue
    return found


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
            ROOT / "bin",
            ROOT / "app",
            home / "Desktop",
            home / "Desktop" / "Afrad2_work",
            home / "Documents",
            home / "Downloads",
            home / "OneDrive" / "Desktop",
            home / "OneDrive" / "Desktop" / "Afrad2_work",
            LEGACY_DIR,
            Path(r"C:\Afrad"),
            Path(r"C:\AFRAD"),
            Path(r"D:\Afrad2_work"),
            Path(r"D:\Afrad"),
        ]
    )
    if os.name == "nt":
        for drive in ready_windows_drives():
            candidates.append(drive)
            candidates.append(drive / "Afrad2_work")
            candidates.append(drive / "Afrad")
    else:
        candidates.extend([Path("/mnt/h"), Path("/mnt/c/Afrad2_work")])
    return _unique_existing(candidates)


def load_cfg() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    if CFG_PATH.exists():
        try:
            cfg.read(CFG_PATH, encoding="utf-8")
        except Exception:
            try:
                cfg.read(CFG_PATH, encoding="cp1256")
            except Exception:
                pass
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
    try:
        return path if path.exists() else None
    except Exception:
        return None


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
    for root in (ROOT, ROOT / "data"):
        if not root.exists():
            continue
        for pat in ("*.xlsx", "*.xls"):
            try:
                hits = sorted(root.glob(pat))
            except Exception:
                continue
            for hit in hits:
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
        try:
            if side.exists():
                files.append(side)
        except Exception:
            continue
    return files


def _copy_if_needed(src: Path, dest: Path) -> None:
    try:
        if not src.exists():
            return
        dest.parent.mkdir(parents=True, exist_ok=True)
        if dest.exists():
            try:
                if src.resolve() == dest.resolve():
                    return
            except Exception:
                pass
        shutil.copy2(src, dest)
    except Exception:
        pass


def ensure_local_mfile() -> Optional[Path]:
    """نسخة عمل محلية من MFILE حتى لا يعتمد البرنامج على H:."""
    local = ROOT / "MFILE.DBF"
    if local.exists():
        return local
    for src in (
        ROOT / "MFILE_updated.DBF",
        ROOT / "data" / "MFILE.DBF",
        ROOT / "data" / "MFILE_updated.DBF",
    ):
        if src.exists():
            try:
                shutil.copy2(src, local)
                for side in companion_files(src)[1:]:
                    shutil.copy2(side, ROOT / side.name)
                return local
            except Exception:
                return src
    found = find_mfile_dbf()
    if not found:
        return None
    try:
        if found.resolve() == local.resolve():
            return local
    except Exception:
        pass
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
        try:
            matches = list(root.glob(pat))
        except Exception:
            continue
        for path in matches:
            try:
                os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
            except Exception:
                pass


def ensure_drive_letter(letter: str, target: Path) -> str:
    """
    ربط حرف قرص قديم (غالباً H:) بالمجلد الحالي.
    لا نفحص أقراصاً أخرى حتى لا تظهر رسالة 'أدخل قرصاً'.
    """
    if os.name != "nt":
        return "not-windows"
    silence_windows_drive_errors()
    drive = f"{letter}:"
    drive_root = Path(f"{drive}\\")
    try:
        result = subprocess.run(
            ["subst", drive, str(target)],
            capture_output=True,
            text=True,
            check=False,
        )
    except Exception:
        result = None
    try:
        if drive_root.exists():
            seed_drive_with_tables(drive_root, target)
            if result is not None and result.returncode == 0:
                return "mapped"
            return "exists"
    except Exception:
        pass
    return "failed"


def seed_drive_with_tables(drive_root: Path, target: Path) -> None:
    for name in ("MFILE.DBF", "MFILE.CDX", "MFILE.FPT", "config.fpw"):
        src = target / name
        dest = drive_root / name
        try:
            if src.exists() and not dest.exists():
                if src.resolve() != dest.resolve():
                    shutil.copy2(src, dest)
        except Exception:
            continue


def ensure_legacy_folder(target: Path) -> str:
    """إنشاء C:\\Afrad2_work كوصلة أو نسخة لأن البرنامج القديم قد يثبّت المسار."""
    if os.name != "nt":
        return "not-windows"
    silence_windows_drive_errors()
    try:
        if target.resolve() == LEGACY_DIR.resolve():
            return "same"
    except Exception:
        if str(target).lower().rstrip("\\/") == str(LEGACY_DIR).lower():
            return "same"
    if not LEGACY_DIR.exists():
        try:
            subprocess.run(
                ["cmd", "/c", "mklink", "/J", str(LEGACY_DIR), str(target)],
                capture_output=True,
                text=True,
                check=False,
            )
        except Exception:
            pass
    if not LEGACY_DIR.exists():
        try:
            LEGACY_DIR.mkdir(parents=True, exist_ok=True)
        except Exception:
            return "failed"
    try:
        if LEGACY_DIR.resolve() == target.resolve():
            return "linked"
    except Exception:
        pass
    for pat in ("*.DBF", "*.CDX", "*.FPT", "*.APP", "*.DLL", "*.EXE", "config.fpw"):
        try:
            for src in target.glob(pat):
                _copy_if_needed(src, LEGACY_DIR / src.name)
        except Exception:
            continue
    try:
        (LEGACY_DIR / "TEMP").mkdir(exist_ok=True)
    except Exception:
        pass
    return "copied"


def write_config_fpw(app_dir: Path) -> Path:
    temp = app_dir / "TEMP"
    try:
        temp.mkdir(exist_ok=True)
    except Exception:
        pass
    content = (
        "DEFAULT=.\n"
        "PATH=.;.\\data;H:\\;C:\\Afrad2_work\n"
        "RESOURCE=OFF\n"
        "HELP=OFF\n"
        "SCREEN=ON\n"
        "TMPFILES=.\\TEMP\n"
        "SORTWORK=.\\TEMP\n"
        "EDITWORK=.\\TEMP\n"
        "MVCOUNT=8192\n"
        "CODEPAGE=1256\n"
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
    try:
        if path.parent.resolve() == ROOT:
            score += 5
    except Exception:
        if path.parent == ROOT:
            score += 5
    try:
        size = path.stat().st_size
        if size > 500_000:
            score += 3
        if size > 2_000_000:
            score += 3
    except Exception:
        pass
    return score


def find_personnel_exe() -> Optional[Path]:
    cached = cfg_path("personnel_exe")
    if cached and cached.suffix.lower() == ".exe" and cached.name.lower() not in EXE_SKIP:
        return cached
    found: List[Tuple[int, Path]] = []
    search_dirs = [ROOT, ROOT / "bin", ROOT / "app", LEGACY_DIR]
    search_dirs.extend(iter_search_roots()[:16])
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
    found.sort(key=lambda x: (x[0], _safe_size(x[1])), reverse=True)
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
        if _same_dir(item[1].parent, root_res) and item[1].name.lower() not in EXE_SKIP
    ]
    if len(local) == 1:
        best = local[0][1]
        save_cfg_value("personnel_exe", best)
        return best
    if local:
        best = max(local, key=lambda x: _safe_size(x[1]))[1]
        save_cfg_value("personnel_exe", best)
        return best
    return None


def _safe_size(path: Path) -> int:
    try:
        return path.stat().st_size
    except Exception:
        return 0


def _same_dir(a: Path, b: Path) -> bool:
    try:
        return a.resolve() == b.resolve()
    except Exception:
        return str(a).lower() == str(b).lower()


def copy_vfp_runtime(app_dir: Path) -> List[Path]:
    copied: List[Path] = []
    search = [
        Path(r"C:\Program Files (x86)\Common Files\Microsoft Shared\VFP"),
        Path(r"C:\Program Files\Common Files\Microsoft Shared\VFP"),
        Path(os.environ.get("WINDIR", r"C:\Windows")) / "SysWOW64",
        Path(os.environ.get("WINDIR", r"C:\Windows")) / "System32",
        LEGACY_DIR,
        ROOT,
    ]
    for name in VFP_RUNTIME_FILES:
        dest = app_dir / name
        if dest.exists():
            continue
        for folder in search:
            src = folder / name
            try:
                if not src.exists():
                    alt = folder / name.lower()
                    src = alt if alt.exists() else src
                if src.exists():
                    try:
                        if src.resolve() == dest.resolve():
                            break
                    except Exception:
                        pass
                    try:
                        shutil.copy2(src, dest)
                        copied.append(dest)
                        break
                    except Exception:
                        continue
            except Exception:
                continue
    return copied


def launch_personnel(exe: Optional[Path] = None) -> int:
    exe = exe or find_personnel_exe()
    if exe is None or not exe.exists():
        print("لم يتم العثور على برنامج نظام الأفراد (ملف EXE).")
        print("انسخ مجلد البرنامج كاملاً من الحاسبة التي يعمل عليها إلى هذا المجلد:")
        print(f"  {ROOT}")
        print("الملفات المهمة: البرنامج EXE + VFP*.DLL + ملفات DBF/CDX")
        return 0
    app_dir = exe.parent
    make_data_writable(app_dir)
    make_data_writable(ROOT)
    make_data_writable(ROOT / "data")
    ensure_local_mfile()
    write_config_fpw(app_dir)
    mapped = ensure_drive_letter("H", app_dir)
    ensure_legacy_folder(app_dir)
    copy_vfp_runtime(app_dir)
    env = os.environ.copy()
    env["PATH"] = str(app_dir) + os.pathsep + env.get("PATH", "")
    print(f"تشغيل نظام الأفراد: {exe}")
    print(f"المجلد: {app_dir}")
    if mapped == "mapped":
        print("تم ربط القرص H: بهذا المجلد حتى تعمل المسارات القديمة.")
    try:
        subprocess.Popen([str(exe)], cwd=str(app_dir), env=env)
        return 0
    except OSError as exc:
        print("تعذر فتح البرنامج:", exc)
        print("انسخ مكتبات FoxPro (VFP9R.DLL أو VFP6R.DLL) من الحاسبة التي يعمل عليها البرنامج")
        print("وضعها في نفس مجلد EXE ثم أعد تشغيل START.bat")
        return 0


def prepare_environment() -> Dict[str, str]:
    """تحضير الجهاز الحالي بدون إظهار أخطاء للمستخدم."""
    silence_windows_drive_errors()
    status = {
        "root": str(ROOT),
        "mfile": "",
        "excel": "",
        "exe": "",
        "hdrive": "unknown",
        "legacy": "unknown",
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
        write_config_fpw(ROOT)
        status["hdrive"] = ensure_drive_letter("H", ROOT)
        status["legacy"] = ensure_legacy_folder(ROOT)
        exe = find_personnel_exe()
        if exe:
            status["exe"] = str(exe)
            write_config_fpw(exe.parent)
            copy_vfp_runtime(exe.parent)
            if exe.parent.resolve() != ROOT.resolve():
                ensure_drive_letter("H", exe.parent)
                ensure_legacy_folder(exe.parent)
    except Exception as exc:
        status["note"] = str(exc)
    return status


def default_update_paths() -> Tuple[Path, Path]:
    excel = find_excel() or (ROOT / "data" / "New Microsoft Excel Worksheet.xlsx")
    dbfp = ensure_local_mfile() or find_mfile_dbf() or (ROOT / "MFILE.DBF")
    return excel, dbfp
