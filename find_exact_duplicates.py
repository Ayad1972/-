# -*- coding: utf-8 -*-
"""
حصر الملفات والمجلدات المتطابقة 100% ثم الإبقاء على الأحدث فقط.

الوضع الافتراضي: تقرير فقط (بدون حذف/نقل).
للتطبيق الفعلي: أضف --apply (ينقل المكررات إلى مجلد حجر، لا يحذف نهائياً).

أمثلة:
  python find_exact_duplicates.py
  python find_exact_duplicates.py --roots "%USERPROFILE%\\Desktop" "%USERPROFILE%\\Downloads"
  python find_exact_duplicates.py --apply
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import os
import shutil
import sys
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple


SKIP_DIR_NAMES = {
    "$recycle.bin",
    "system volume information",
    ".git",
    ".svn",
    "__pycache__",
    "node_modules",
    "windows",
    "program files",
    "program files (x86)",
    "programdata",
}

SKIP_FILE_PREFIXES = ("~$",)  # Excel temp locks


@dataclass
class FileInfo:
    path: Path
    size: int
    mtime: float
    digest: str


def default_roots() -> List[Path]:
    home = Path.home()
    candidates = [
        home / "Desktop",
        home / "OneDrive" / "Desktop",
        home / "Downloads",
        Path(r"H:\\"),
    ]
    # Arabic Windows may use localized names; also try environment
    desktop = os.environ.get("USERPROFILE")
    if desktop:
        candidates.insert(0, Path(desktop) / "Desktop")
        candidates.insert(1, Path(desktop) / "Downloads")
    out: List[Path] = []
    seen = set()
    for p in candidates:
        try:
            rp = p.resolve()
        except Exception:
            rp = p
        key = str(rp).lower()
        if key in seen:
            continue
        if p.exists() and p.is_dir():
            seen.add(key)
            out.append(p)
    return out


def should_skip_dir(path: Path) -> bool:
    name = path.name.lower()
    if name in SKIP_DIR_NAMES:
        return True
    if name.startswith(".") and name not in {".", ".."}:
        # keep normal folders; skip hidden tooling mostly already listed
        if name in {".cursor", ".vscode", ".idea"}:
            return True
    return False


def iter_files(roots: Sequence[Path]) -> Iterable[Path]:
    for root in roots:
        if not root.exists():
            continue
        for dirpath, dirnames, filenames in os.walk(root):
            # prune in-place
            dirnames[:] = [
                d for d in dirnames
                if not should_skip_dir(Path(dirpath) / d)
            ]
            for fn in filenames:
                if fn.startswith(SKIP_FILE_PREFIXES):
                    continue
                p = Path(dirpath) / fn
                try:
                    if p.is_file() and not p.is_symlink():
                        yield p
                except OSError:
                    continue


def sha256_file(path: Path, chunk: int = 1024 * 1024) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            block = f.read(chunk)
            if not block:
                break
            h.update(block)
    return h.hexdigest()


def collect_file_infos(roots: Sequence[Path]) -> List[FileInfo]:
    # Phase A: group by size first (fast reject)
    by_size: Dict[int, List[Path]] = defaultdict(list)
    for p in iter_files(roots):
        try:
            st = p.stat()
        except OSError:
            continue
        if st.st_size <= 0:
            continue
        by_size[st.st_size].append(p)

    infos: List[FileInfo] = []
    # Only hash sizes that appear more than once (potential duplicates)
    # BUT for folder signatures we need ALL file hashes under compared folders.
    # Strategy:
    # 1) Hash all files that share size with another file (for file-dup report)
    # 2) Separately compute folder signatures only for sibling-like folder candidates
    for size, paths in by_size.items():
        if len(paths) < 2:
            continue
        for p in paths:
            try:
                st = p.stat()
                digest = sha256_file(p)
                infos.append(FileInfo(p, st.st_size, st.st_mtime, digest))
            except OSError:
                continue
    return infos


def newest_first(items: Sequence[FileInfo]) -> List[FileInfo]:
    # Keep the newest mtime; tie-breaker: longer path last then lexical last
    return sorted(items, key=lambda x: (x.mtime, str(x.path).lower()), reverse=True)


def group_identical_files(infos: Sequence[FileInfo]) -> List[List[FileInfo]]:
    by_hash: Dict[str, List[FileInfo]] = defaultdict(list)
    for info in infos:
        by_hash[info.digest].append(info)
    groups = [g for g in by_hash.values() if len(g) >= 2]
    # stable order by size desc
    groups.sort(key=lambda g: (-g[0].size, g[0].digest))
    return groups


def folder_signature(folder: Path, file_hash_cache: Dict[Path, str]) -> Optional[Tuple[Tuple[str, str, int], ...]]:
    """Signature = sorted tuples of (relpath, sha256, size) for all files under folder."""
    entries: List[Tuple[str, str, int]] = []
    if not folder.is_dir():
        return None
    for dirpath, dirnames, filenames in os.walk(folder):
        dirnames[:] = [d for d in dirnames if not should_skip_dir(Path(dirpath) / d)]
        for fn in filenames:
            if fn.startswith(SKIP_FILE_PREFIXES):
                continue
            p = Path(dirpath) / fn
            try:
                if not p.is_file() or p.is_symlink():
                    continue
                st = p.stat()
                if st.st_size <= 0:
                    # include empty files in signature for exactness
                    digest = hashlib.sha256(b"").hexdigest()
                else:
                    if p not in file_hash_cache:
                        file_hash_cache[p] = sha256_file(p)
                    digest = file_hash_cache[p]
                rel = p.relative_to(folder).as_posix().lower()
                entries.append((rel, digest, st.st_size))
            except OSError:
                continue
    if not entries:
        return None  # ignore empty folders
    return tuple(sorted(entries))


def collect_folders(roots: Sequence[Path], max_depth: int = 4) -> List[Path]:
    folders: List[Path] = []
    for root in roots:
        root = root.resolve()
        for dirpath, dirnames, _filenames in os.walk(root):
            p = Path(dirpath)
            try:
                depth = len(p.relative_to(root).parts)
            except ValueError:
                depth = 0
            # prune deep trees
            if depth >= max_depth:
                dirnames[:] = []
            dirnames[:] = [d for d in dirnames if not should_skip_dir(p / d)]
            if p == root:
                continue
            if depth <= max_depth:
                folders.append(p)
    return folders


def is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except Exception:
        return False


def filter_top_level_folder_groups(groups: Sequence[Sequence[Path]]) -> List[List[Path]]:
    """إذا تطابق مجلدان أبوان، لا تكرّر الإبلاغ عن كل مجلد فرعي داخلهما."""
    # اجمع كل المجلدات التي ستُعتبر مكررة (غير الأحدث) في مجموعات الآباء
    covered_roots: List[Path] = []
    tentative: List[List[Path]] = []
    # الأكبر أولاً (أقل عمقاً أولاً)
    ranked = sorted(
        [list(g) for g in groups if len(g) >= 2],
        key=lambda g: (min(len(p.parts) for p in g), -max(folder_mtime(p) for p in g)),
    )
    for group in ranked:
        ordered = sorted(group, key=lambda p: (folder_mtime(p), str(p).lower()), reverse=True)
        # تخطّي المجموعة إذا كانت كل عناصرها داخل جذور مغطاة مسبقاً بنفس العلاقات
        if all(any(is_relative_to(p, root) and p != root for root in covered_roots) for p in ordered):
            continue
        tentative.append(ordered)
        # غطِّ كل نسخ المجموعة (بما فيها التي ستُبقى) لمنع تكرار الأبناء
        covered_roots.extend(ordered)
    return tentative


def group_identical_folders(roots: Sequence[Path], max_depth: int = 4) -> List[List[Path]]:
    cache: Dict[Path, str] = {}
    by_sig: Dict[Tuple[Tuple[str, str, int], ...], List[Path]] = defaultdict(list)
    folders = collect_folders(roots, max_depth=max_depth)
    total = len(folders)
    for i, folder in enumerate(folders, 1):
        if i % 50 == 0 or i == total:
            print(f"  مسح المجلدات: {i}/{total}", flush=True)
        sig = folder_signature(folder, cache)
        if not sig:
            continue
        by_sig[sig].append(folder)

    groups: List[List[Path]] = []
    for folders_g in by_sig.values():
        if len(folders_g) < 2:
            continue
        groups.append(
            sorted(
                folders_g,
                key=lambda p: (folder_mtime(p), str(p).lower()),
                reverse=True,
            )
        )
    return filter_top_level_folder_groups(groups)


def folder_mtime(path: Path) -> float:
    try:
        return path.stat().st_mtime
    except OSError:
        return 0.0


def human_size(n: int) -> str:
    units = ["B", "KB", "MB", "GB", "TB"]
    f = float(n)
    for u in units:
        if f < 1024 or u == units[-1]:
            return f"{f:.1f} {u}"
        f /= 1024
    return f"{n} B"


def quarantine_base(roots: Sequence[Path]) -> Path:
    # Prefer Desktop for quarantine visibility
    for r in roots:
        if r.name.lower() == "desktop":
            return r / "_DUPLICATES_QUARANTINE"
    return Path.home() / "Desktop" / "_DUPLICATES_QUARANTINE"


def unique_dest(dest: Path) -> Path:
    if not dest.exists():
        return dest
    stem, suffix = dest.stem, dest.suffix
    parent = dest.parent
    i = 2
    while True:
        cand = parent / f"{stem}__{i}{suffix}"
        if not cand.exists():
            return cand
        i += 1


def move_to_quarantine(src: Path, qroot: Path, root_hint: Path) -> Path:
    try:
        rel = src.relative_to(root_hint)
    except Exception:
        rel = Path(src.name)
    dest = qroot / rel
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest = unique_dest(dest)
    shutil.move(str(src), str(dest))
    return dest


def write_reports(
    out_dir: Path,
    file_groups: Sequence[Sequence[FileInfo]],
    folder_groups: Sequence[Sequence[Path]],
) -> Tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    txt_path = out_dir / f"duplicates_report_{stamp}.txt"
    csv_path = out_dir / f"duplicates_report_{stamp}.csv"

    keep_folders = [g[0] for g in folder_groups if g]

    reclaimable = 0
    with txt_path.open("w", encoding="utf-8") as f:
        f.write("تقرير الملفات والمجلدات المتطابقة 100%\n")
        f.write(f"التاريخ: {datetime.now().isoformat(timespec='seconds')}\n")
        f.write("القاعدة: الإبقاء على الأحدث/الأخير، ونقل الباقي عند --apply إلى الحجر\n")
        f.write("=" * 70 + "\n\n")

        f.write(f"أولاً: مجلدات متطابقة 100% — عدد المجموعات: {len(folder_groups)}\n\n")
        for idx, group in enumerate(folder_groups, 1):
            ordered = list(group)
            keep = ordered[0]
            drop = ordered[1:]
            f.write(f"[مجلد مجموعة {idx}] عدد النسخ={len(ordered)}\n")
            f.write(f"  يُبقى (الأخير/الأحدث): {keep}\n")
            f.write(f"             التاريخ: {datetime.fromtimestamp(folder_mtime(keep))}\n")
            for d in drop:
                f.write(f"  مكرر (يُنقل): {d}\n")
                f.write(f"             التاريخ: {datetime.fromtimestamp(folder_mtime(d))}\n")
            f.write("\n")

        f.write("=" * 70 + "\n\n")
        f.write(f"ثانياً: ملفات متطابقة 100% — عدد المجموعات: {len(file_groups)}\n\n")
        for idx, group in enumerate(file_groups, 1):
            keep = choose_file_keep(group, keep_folders)
            drop = [x for x in group if x.path != keep.path]
            # تقديري: لا نحسب ملفات داخل المجلدات المكررة التي ستُنقل كاملة
            for d in drop:
                reclaimable += d.size
            f.write(f"[ملف مجموعة {idx}] الحجم={human_size(keep.size)}  hash={keep.digest[:12]}...\n")
            f.write(f"  يُبقى: {keep.path}\n")
            f.write(f"        التاريخ: {datetime.fromtimestamp(keep.mtime)}\n")
            for d in drop:
                f.write(f"  مكرر (يُنقل إن لم يكن داخل مجلد مُبقى): {d.path}\n")
                f.write(f"        التاريخ: {datetime.fromtimestamp(d.mtime)}\n")
            f.write("\n")

        f.write("=" * 70 + "\n")
        f.write(f"حجم تقديري يمكن تحريره من الملفات المكررة: {human_size(reclaimable)}\n")

    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["type", "group", "action", "path", "mtime", "size", "hash"])
        for idx, group in enumerate(folder_groups, 1):
            for i, item in enumerate(group):
                w.writerow([
                    "folder",
                    idx,
                    "KEEP" if i == 0 else "MOVE",
                    str(item),
                    datetime.fromtimestamp(folder_mtime(item)).isoformat(timespec="seconds"),
                    "",
                    "",
                ])
        for idx, group in enumerate(file_groups, 1):
            keep = choose_file_keep(group, keep_folders)
            for item in group:
                w.writerow([
                    "file",
                    idx,
                    "KEEP" if item.path == keep.path else "MOVE",
                    str(item.path),
                    datetime.fromtimestamp(item.mtime).isoformat(timespec="seconds"),
                    item.size,
                    item.digest,
                ])

    return txt_path, csv_path


def choose_file_keep(group: Sequence[FileInfo], keep_folders: Sequence[Path]) -> FileInfo:
    """يفضّل نسخة داخل مجلد مُبقى، وإلا الأحدث."""
    ordered = newest_first(list(group))
    for item in ordered:
        if any(is_relative_to(item.path, kf) for kf in keep_folders):
            return item
    return ordered[0]


def apply_moves(
    file_groups: Sequence[Sequence[FileInfo]],
    folder_groups: Sequence[Sequence[Path]],
    roots: Sequence[Path],
) -> Tuple[int, int]:
    qroot = quarantine_base(roots)
    qroot.mkdir(parents=True, exist_ok=True)
    moved_files = 0
    moved_folders = 0

    keep_folders: List[Path] = []
    moved_folder_paths: List[Path] = []

    # Folders first (أبقى الأحدث، انقل الباقي كاملاً)
    for group in folder_groups:
        ordered = sorted(group, key=lambda p: (folder_mtime(p), str(p).lower()), reverse=True)
        keep_folders.append(ordered[0])
        for folder in ordered[1:]:
            if not folder.exists():
                continue
            root_hint = roots[0]
            for r in roots:
                try:
                    folder.relative_to(r)
                    root_hint = r
                    break
                except Exception:
                    continue
            try:
                dest = move_to_quarantine(folder, qroot / "folders", root_hint)
                print(f"نُقل مجلد: {folder} -> {dest}")
                moved_folders += 1
                moved_folder_paths.append(folder)
            except Exception as exc:
                print(f"فشل نقل مجلد: {folder} ({exc})")

    def under_moved_folder(path: Path) -> bool:
        return any(is_relative_to(path, mf) for mf in moved_folder_paths)

    def under_keep_folder(path: Path) -> bool:
        return any(is_relative_to(path, kf) for kf in keep_folders)

    # ثم الملفات المكررة خارج المجلدات التي نُقلت، دون تفريغ المجلدات المُبقاة
    for group in file_groups:
        keep = choose_file_keep(group, keep_folders)
        for item in group:
            if item.path == keep.path:
                continue
            if not item.path.exists():
                continue
            if under_moved_folder(item.path):
                continue
            if under_keep_folder(item.path):
                # لا نُفرّغ المجلد الذي أبقيناه كآخر نسخة متطابقة
                continue
            root_hint = roots[0]
            for r in roots:
                try:
                    item.path.relative_to(r)
                    root_hint = r
                    break
                except Exception:
                    continue
            try:
                dest = move_to_quarantine(item.path, qroot / "files", root_hint)
                print(f"نُقل ملف: {item.path} -> {dest}")
                moved_files += 1
            except Exception as exc:
                print(f"فشل نقل ملف: {item.path} ({exc})")

    print(f"مجلد الحجر: {qroot}")
    return moved_files, moved_folders


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="حصر الملفات/المجلدات المتطابقة 100% والإبقاء على الأحدث")
    p.add_argument(
        "--roots",
        nargs="*",
        default=None,
        help="المسارات للفحص (افتراضي: Desktop و Downloads إن وُجدت)",
    )
    p.add_argument(
        "--apply",
        action="store_true",
        help="تطبيق النقل إلى مجلد الحجر (بدون هذا الخيار: تقرير فقط)",
    )
    p.add_argument(
        "--max-folder-depth",
        type=int,
        default=4,
        help="أقصى عمق لمقارنة المجلدات داخل كل جذر (افتراضي 4)",
    )
    p.add_argument(
        "--report-dir",
        default=None,
        help="مجلد حفظ التقرير (افتراضي: سطح المكتب أو مجلد المشروع)",
    )
    return p.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    if args.roots:
        roots = [Path(os.path.expandvars(os.path.expanduser(r))) for r in args.roots]
        roots = [r for r in roots if r.exists()]
    else:
        roots = default_roots()

    if not roots:
        print("لا توجد مسارات للفحص. مرّر --roots مثل:")
        print('  python find_exact_duplicates.py --roots "%USERPROFILE%\\Desktop"')
        return 1

    print("=" * 60)
    print("حصر التطابق 100% — تقرير أولاً")
    print("=" * 60)
    for r in roots:
        print(f"مسار الفحص: {r}")
    print(f"وضع التطبيق: {'نعم — نقل للمكررات' if args.apply else 'لا — تقرير فقط'}")
    print()

    print("1) فحص الملفات المتطابقة...")
    file_infos = collect_file_infos(roots)
    file_groups = group_identical_files(file_infos)
    print(f"   مجموعات ملفات مكررة: {len(file_groups)}")

    print("2) فحص المجلدات المتطابقة...")
    folder_groups = group_identical_folders(roots, max_depth=args.max_folder_depth)
    print(f"   مجموعات مجلدات مكررة: {len(folder_groups)}")

    report_dir = Path(args.report_dir) if args.report_dir else None
    if report_dir is None:
        for r in roots:
            if r.name.lower() == "desktop":
                report_dir = r
                break
        if report_dir is None:
            report_dir = Path.cwd()

    txt_path, csv_path = write_reports(report_dir, file_groups, folder_groups)
    print()
    print("تم حفظ التقرير:")
    print(f"  {txt_path}")
    print(f"  {csv_path}")

    if not args.apply:
        print()
        print("لم يتم نقل أي شيء. راجع التقرير أولاً.")
        print("إذا كان صحيحاً، شغّل نفس الأمر مع --apply")
        return 0

    print()
    print("3) نقل المكررات إلى الحجر والإبقاء على الأحدث فقط...")
    moved_files, moved_folders = apply_moves(file_groups, folder_groups, roots)
    print(f"تم نقل ملفات: {moved_files}")
    print(f"تم نقل مجلدات: {moved_folders}")
    print("راجع مجلد _DUPLICATES_QUARANTINE على سطح المكتب.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
