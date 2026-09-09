# -*- coding: utf-8 -*-
"""
حصر مجلدات الفلاش التي يحتوي جدول mfile.dbf فيها على رقم موظف محدد.
لا يحتاج مكتبات خارجية.
"""

from __future__ import annotations

import argparse
import os
import struct
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple


TARGET_DEFAULT = "396176"

# ترميز أسماء نظام الفوكس العراقي (مستنتج من سجلات معروفة)
FOX_AR_MAP = {
    0xAB: "ا",
    0xAC: "ب",
    0xE0: "ت",
    0xE1: "ث",
    0xE2: "ج",
    0xE3: "ح",
    0xE4: "خ",
    0xE5: "د",
    0xE6: "ذ",
    0xE7: "ر",
    0xE8: "ز",
    0xE9: "س",
    0xEA: "ش",
    0xEB: "ص",
    0xEC: "ض",
    0xED: "ط",
    0xEE: "ظ",
    0xEF: "ع",
    0xF0: "غ",
    0xF2: "ف",
    0xF3: "ق",
    0xF4: "ك",
    0xF5: "ل",
    0xF6: "م",
    0xF8: "ن",
    0xF9: "ه",
    0xFA: "ة",
    0xFB: "و",
    0xFC: "ى",
    0xFD: "ي",
}

MFILE_NAMES = {"mfile.dbf", "mfile_updated.dbf"}


def decode_fox_ar(raw: bytes) -> str:
    out: List[str] = []
    for b in raw.rstrip(b" \x00"):
        if b == 0x20:
            out.append(" ")
        else:
            out.append(FOX_AR_MAP.get(b, f"[{b:02X}]"))
    return " ".join("".join(out).split())


def parse_dbf_fields(header: bytes, header_len: int) -> List[Tuple[str, str, int, int]]:
    fields: List[Tuple[str, str, int, int]] = []
    off = 32
    while off + 32 <= header_len:
        raw = header[off : off + 32]
        if raw[0] in (0x0D, 0x00):
            break
        name = raw[0:11].split(b"\x00", 1)[0].decode("ascii", "replace").strip()
        ftype = chr(raw[11]) if raw[11] else "C"
        flen = raw[16]
        fdec = raw[17]
        fields.append((name.upper(), ftype, flen, fdec))
        off += 32
    return fields


def iter_dbf_records(path: Path):
    data = path.read_bytes()
    if len(data) < 32 or data[0] not in (0x03, 0x30, 0x31, 0x83, 0x8B, 0xF5):
        raise ValueError("ليس ملف DBF صالح")
    nrecords = struct.unpack_from("<I", data, 4)[0]
    header_len = struct.unpack_from("<H", data, 8)[0]
    rec_len = struct.unpack_from("<H", data, 10)[0]
    if header_len < 33 or rec_len < 2:
        raise ValueError("هيدر DBF غير مكتمل")
    fields = parse_dbf_fields(data[:header_len], header_len)
    for i in range(nrecords):
        start = header_len + i * rec_len
        rec = data[start : start + rec_len]
        if len(rec) < rec_len:
            break
        yield rec[0:1], fields, rec[1:]


def record_to_dict(fields: Sequence[Tuple[str, str, int, int]], body: bytes) -> Dict[str, bytes]:
    out: Dict[str, bytes] = {}
    pos = 0
    for name, _ftype, flen, _fdec in fields:
        out[name] = body[pos : pos + flen]
        pos += flen
        if pos >= len(body):
            break
    return out


def pno_matches(value: bytes, target: str) -> bool:
    text = value.decode("ascii", "ignore").strip()
    if not text:
        return False
    if text.isdigit():
        return str(int(text)) == str(int(target)) if target.isdigit() else text == target
    return text == target


def find_pno_in_dbf(path: Path, target: str) -> List[Dict[str, str]]:
    hits: List[Dict[str, str]] = []
    for deleted, fields, body in iter_dbf_records(path):
        if deleted == b"*":
            continue
        row = record_to_dict(fields, body)
        pno_key = None
        for key in ("PNO", "EMPNO", "EMP_NO", "NO", "NUM", "ID"):
            if key in row:
                pno_key = key
                break
        if pno_key is None:
            # أول حقل نصي/رقمي قصير غالباً رقم الموظف
            for name, _t, flen, _d in fields:
                if 4 <= flen <= 10 and name in row:
                    pno_key = name
                    break
        if pno_key is None:
            continue
        if not pno_matches(row[pno_key], target):
            continue
        name_raw = b""
        for key in ("NAME", "ENAME", "EMPNAME", "FULLNAME"):
            if key in row:
                name_raw = row[key]
                break
        dd = row.get("DD", b"").decode("ascii", "ignore").strip()
        mm = row.get("MM", b"").decode("ascii", "ignore").strip()
        yy = row.get("YY", b"").decode("ascii", "ignore").strip()
        hits.append(
            {
                "pno": row[pno_key].decode("ascii", "ignore").strip(),
                "name": decode_fox_ar(name_raw) if name_raw else "",
                "dd": dd,
                "mm": mm,
                "yy": yy,
                "deleted": "no",
            }
        )
    if hits:
        return hits

    # احتياط: بحث بايتات إذا اختلف اسم الحقل
    raw = path.read_bytes()
    needle = target.encode("ascii")
    if needle in raw:
        return [
            {
                "pno": target,
                "name": "",
                "dd": "",
                "mm": "",
                "yy": "",
                "deleted": "unknown",
            }
        ]
    return []


def is_mfile(path: Path) -> bool:
    return path.is_file() and path.name.lower() in MFILE_NAMES


def iter_mfile_paths(roots: Iterable[Path]) -> Iterable[Path]:
    seen = set()
    for root in roots:
        if not root.exists():
            continue
        try:
            root = root.resolve()
        except OSError:
            pass
        if is_mfile(root):
            key = str(root).lower()
            if key not in seen:
                seen.add(key)
                yield root
            continue
        if not root.is_dir():
            continue
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if d not in {".git", "node_modules"}]
            for name in filenames:
                if name.lower() in MFILE_NAMES:
                    p = Path(dirpath) / name
                    key = str(p).lower()
                    if key not in seen:
                        seen.add(key)
                        yield p


def windows_flash_roots() -> List[Path]:
    roots: List[Path] = []
    if os.name == "nt":
        # H: هو حرف الفلاش المعتاد في هذا النظام
        for letter in ["H"] + [chr(c) for c in range(ord("D"), ord("Z") + 1) if chr(c) != "H"]:
            drive = Path(f"{letter}:\\")
            try:
                if drive.exists():
                    roots.append(drive)
            except OSError:
                continue
    else:
        for p in (Path("/mnt/h"), Path("/media/H"), Path("/media/ubuntu")):
            if p.exists():
                roots.append(p)
    return roots


def default_roots(extra: Sequence[str]) -> List[Path]:
    here = Path(__file__).resolve().parent
    roots: List[Path] = []
    for item in extra:
        roots.append(Path(item))
    if extra:
        # إذا حدّد المستخدم المسارات فلا نخلط معها مسارات أخرى
        pass
    else:
        roots.extend(windows_flash_roots())
        roots.extend(
            [
                here,
                here / "data",
                Path.cwd(),
            ]
        )
    # أزل التكرار مع الإبقاء على الترتيب
    uniq: List[Path] = []
    seen = set()
    for r in roots:
        key = str(r).lower()
        if key not in seen:
            seen.add(key)
            uniq.append(r)
    return uniq


def format_hit_line(folder: Path, dbf: Path, hit: Dict[str, str]) -> str:
    svc = ""
    if hit["dd"] or hit["mm"] or hit["yy"]:
        svc = f" | DD/MM/YY=({hit['dd']}, {hit['mm']}, {hit['yy']})"
    name = hit["name"] or "-"
    return f"{folder} | {dbf.name} | PNO={hit['pno']} | {name}{svc}"


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="حصر مجلدات mfile.dbf التي فيها رقم موظف")
    parser.add_argument("--pno", default=TARGET_DEFAULT, help="رقم الموظف")
    parser.add_argument("--root", action="append", default=[], help="مسار للفحص (يمكن تكراره)")
    parser.add_argument("--out", default="", help="ملف تقرير اختياري")
    args = parser.parse_args(argv)

    target = str(args.pno).strip()
    if not target:
        print("رقم الموظف فارغ")
        return 2

    roots = default_roots(args.root)
    print("رقم الموظف:", target)
    print("مسارات الفحص:")
    for r in roots:
        mark = "موجود" if r.exists() else "غير موجود"
        print(f"  {r}  [{mark}]")
    print("-" * 60)

    scanned = 0
    matched_folders: List[str] = []
    lines: List[str] = []
    errors: List[str] = []

    for dbf in iter_mfile_paths(roots):
        scanned += 1
        try:
            hits = find_pno_in_dbf(dbf, target)
        except Exception as exc:
            errors.append(f"{dbf}: {exc}")
            continue
        if not hits:
            continue
        folder = dbf.parent
        folder_s = str(folder)
        if folder_s not in matched_folders:
            matched_folders.append(folder_s)
        for hit in hits:
            line = format_hit_line(folder, dbf, hit)
            lines.append(line)
            print("وجد:", line)

    print("-" * 60)
    print("ملفات mfile المفحوصة:", scanned)
    print("المجلدات التي فيها الرقم:", len(matched_folders))
    if matched_folders:
        print("المجلدات:")
        for folder in matched_folders:
            print(folder)
    else:
        print("لا يوجد هذا الرقم في جداول mfile داخل المسارات المتاحة هنا.")
        print("إذا كان الفلاش على ويندوز، شغّل الملف:")
        print("find_pno_in_flash.bat")
    if errors:
        print("أخطاء قراءة:")
        for err in errors:
            print(" ", err)

    if args.out:
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        report = [
            f"PNO={target}",
            f"scanned={scanned}",
            f"folders={len(matched_folders)}",
            "",
            "FOLDERS:",
            *matched_folders,
            "",
            "DETAILS:",
            *lines,
            "",
            "ERRORS:",
            *errors,
        ]
        out.write_text("\n".join(report) + "\n", encoding="utf-8")
        print("التقرير:", out)

    return 0 if matched_folders else 1


if __name__ == "__main__":
    sys.exit(main())
