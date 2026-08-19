# -*- coding: utf-8 -*-
"""قائمة تشغيل نظام الأفراد على أي حاسبة ويندوز."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from afrad_portable import (
    ROOT,
    FF_NAMES,
    MM_NAMES,
    SS_NAMES,
    default_update_paths,
    find_named,
    find_personnel_exe,
    launch_personnel,
    prepare_environment,
    setup_stdio,
)


def run_py(script: str, args: list[str] | None = None) -> int:
    cmd = [sys.executable, str(ROOT / script)]
    if args:
        cmd.extend(args)
    try:
        return subprocess.call(cmd, cwd=str(ROOT))
    except Exception as exc:
        print("تعذر تشغيل الأداة:", exc)
        return 1


def ensure_libs() -> None:
    missing = []
    for mod, pkg in (("openpyxl", "openpyxl"), ("dbfread", "dbfread"), ("dbf", "dbf")):
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)
    try:
        import xlrd  # noqa: F401
    except ImportError:
        missing.append("xlrd==1.2.0")
    if not missing:
        return
    print("تثبيت المكتبات المطلوبة مرة واحدة...")
    cmd = [sys.executable, "-m", "pip", "install", "-q", *missing, "xlwt", "xlutils"]
    try:
        subprocess.check_call(cmd)
    except Exception:
        print("تعذر تثبيت المكتبات تلقائياً. شغّل install_python.bat ثم أعد المحاولة.")


def print_status(status: dict[str, str]) -> None:
    print("=" * 60)
    print("  نظام الأفراد - تشغيل على أي حاسبة")
    print("=" * 60)
    print(f"المجلد: {status.get('root', ROOT)}")
    print(f"MFILE : {status.get('mfile') or 'غير موجود بعد (انسخه إلى هذا المجلد)'}")
    print(f"Excel : {status.get('excel') or 'اختياري - ضعه في المجلد أو data'}")
    print(f"EXE   : {status.get('exe') or 'غير موجود - انسخ برنامج الأفراد إلى هذا المجلد'}")
    hdrive = status.get("hdrive", "")
    if hdrive == "mapped":
        print("القرص H: تم ربطه بهذا المجلد (للتوافق مع الفوكس القديم)")
    elif hdrive == "exists":
        print("القرص H: موجود على هذا الجهاز")
    print("=" * 60)


def cmd_check() -> int:
    status = prepare_environment()
    print_status(status)
    print("CHECK_OK")
    return 0


def show_menu() -> int:
    status = prepare_environment()
    while True:
        print()
        print_status(status)
        print("1) فتح نظام الأفراد")
        print("2) نقل بيانات Excel إلى MFILE.DBF")
        print("3) فحص الملفات فقط")
        print("4) دمج الأرقام حسب الاسم")
        print("5) تثبيت Python والمكتبات")
        print("6) تسريع الحاسوب")
        print("7) نسخ MFILE إلى سطح المكتب")
        print("0) خروج")
        print()
        try:
            choice = input("اختر رقماً: ").strip()
        except EOFError:
            return 0
        if choice == "1":
            launch_personnel()
        elif choice == "2":
            ensure_libs()
            excel, dbfp = default_update_paths()
            run_py(
                "update_mfile_from_excel.py",
                ["--excel", str(excel), "--dbf", str(dbfp)],
            )
        elif choice == "3":
            ensure_libs()
            excel, dbfp = default_update_paths()
            run_py(
                "update_mfile_from_excel.py",
                ["--inspect-only", "--excel", str(excel), "--dbf", str(dbfp)],
            )
        elif choice == "4":
            ensure_libs()
            ss = find_named(SS_NAMES, "ss")
            mm = find_named(MM_NAMES, "mm") or find_named(FF_NAMES, "ff")
            if not ss or not mm:
                print("ضع ss.xls و mm.xlsx (أو ff.xlsx) في هذا المجلد أو في data ثم أعد المحاولة.")
            else:
                out = ss.with_name("ss_updated.xlsx")
                run_py(
                    "merge_mm_into_ss.py",
                    ["--ss", str(ss), "--mm", str(mm), "--out", str(out)],
                )
        elif choice == "5":
            bat = ROOT / "install_python.bat"
            if os_is_windows() and bat.exists():
                subprocess.call(["cmd", "/c", str(bat)], cwd=str(ROOT))
            else:
                ensure_libs()
                print("المكتبات جاهزة.")
        elif choice == "6":
            go = ROOT / "GO.bat"
            if os_is_windows() and go.exists():
                subprocess.call(["cmd", "/c", str(go)], cwd=str(ROOT))
            else:
                print("أداة التسريع مخصصة لويندوز. شغّل GO.bat على الحاسبة.")
        elif choice == "7":
            run_copy_mfile()
        elif choice in {"0", "q", "Q"}:
            print("إلى اللقاء.")
            return 0
        else:
            print("اختيار غير صالح.")
        status = prepare_environment()


def os_is_windows() -> bool:
    return sys.platform.startswith("win")


def run_copy_mfile() -> None:
    from afrad_portable import ensure_local_mfile
    import shutil

    src = ensure_local_mfile()
    if not src:
        print("لا يوجد MFILE.DBF لنسخه.")
        return
    dest = Path.home() / "Desktop" / src.name
    try:
        shutil.copy2(src, dest)
        print("تم النسخ إلى:", dest)
    except Exception as exc:
        fallback = ROOT / src.name
        print("تعذر النسخ إلى سطح المكتب:", exc)
        print("الملف المحلي:", fallback)


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    setup_stdio()
    try:
        if "--check" in argv:
            return cmd_check()
        if "--menu" in argv:
            return show_menu()
        if "--open" in argv or not argv:
            status = prepare_environment()
            print_status(status)
            exe = find_personnel_exe()
            if exe:
                return launch_personnel(exe)
            print()
            print("لم يُفتح برنامج EXE لأن ملف النظام غير موجود في هذا المجلد.")
            print("يمكنك استخدام الأدوات من القائمة، أو نسخ برنامج الأفراد إلى هنا.")
            print()
            return show_menu()
        print("استخدام: python start_afrad.py [--open|--menu|--check]")
        return 0
    except KeyboardInterrupt:
        print("\nتم الإلغاء.")
        return 0
    except Exception as exc:
        print("حدث أمر غير متوقع:", exc)
        print("انسخ مجلد المشروع كاملاً من الحاسبة التي يعمل عليها النظام ثم شغّل START.bat")
        return 0


if __name__ == "__main__":
    sys.exit(main())
