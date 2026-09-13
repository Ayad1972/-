# -*- coding: utf-8 -*-
"""
تفريغ AL082026.DBF دون تغيير هيكله، ثم نقل بيانات Excel إليه
بنفس ترميز العربي الموجود في الجدول.
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import shutil
import struct
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import openpyxl
except ImportError:
    print("المكتبة openpyxl غير مثبتة. نفّذ: pip install openpyxl")
    sys.exit(1)


CODEPAGE_MAP = {
    0x01: "cp437",
    0x02: "cp850",
    0x03: "cp1252",
    0x57: "cp1252",
    0x64: "cp852",
    0x65: "cp866",
    0x66: "cp865",
    0x67: "cp861",
    0x6A: "cp737",
    0x6B: "cp857",
    0x70: "cp720",
    0x7D: "cp1255",
    0x7E: "cp1256",
    0xC8: "cp1250",
    0xC9: "cp1251",
    0xCA: "cp1254",
    0xCB: "cp1253",
}

ARABIC_RE = re.compile(r"[\u0600-\u06FF]")
DIGITS_RE = re.compile(r"\d+")

FIELD_ALIASES: Dict[str, Sequence[str]] = {
    "PNO": (
        "pno", "empno", "emp_no", "no", "num", "number", "id", "code",
        "رقمالموظف", "الرقم", "رقم", "رقموظيفي", "الرقمالوظيفي", "الكود", "كود",
    ),
    "NAME": (
        "name", "ename", "aname", "empname", "fullname",
        "الاسم", "اسم", "اسمالموظف", "الاسمالثلاثي", "الاسمالرباعي",
    ),
    "AMT": (
        "amt", "amount", "allw", "allow", "alawa", "sal", "value", "val", "mony", "money",
        "المبلغ", "مبلغ", "العلاوه", "العلاوة", "علاوه", "علاوة", "المخصص", "القيمه", "القيمة",
    ),
    "ALLW": (
        "allw", "allow", "alawa", "amt", "amount",
        "العلاوه", "العلاوة", "علاوه", "علاوة", "المبلغ",
    ),
    "TYPE": (
        "type", "typ", "kind", "class",
        "النوع", "نوع", "نوعالعلاوه", "نوعالعلاوة", "الصنف", "صنف",
    ),
    "TYP": (
        "typ", "type", "kind",
        "النوع", "نوع", "الصنف",
    ),
    "AMER": (
        "amer", "amr", "order", "ordno", "noamr", "book", "amerno",
        "الامر", "الأمر", "رقمالامر", "رقمالأمر", "رقمالأمر", "امر", "أمر",
    ),
    "DT": (
        "dt", "date", "fdate", "hdate", "adat",
        "التاريخ", "تاريخ", "تاريخالامر", "تاريخالأمر",
    ),
    "DATE": (
        "date", "dt", "fdate",
        "التاريخ", "تاريخ",
    ),
    "NOTE": (
        "note", "notes", "remark", "rem", "obs",
        "ملاحظه", "ملاحظة", "الملاحظات", "ملاحظات",
    ),
    "NOTES": (
        "notes", "note", "remark",
        "الملاحظات", "ملاحظات", "ملاحظة",
    ),
}


@dataclass
class DbfField:
    name: str
    type: str
    length: int
    decimal: int
    offset: int


@dataclass
class DbfMeta:
    path: Path
    header: bytes
    header_len: int
    record_len: int
    record_count: int
    ldid: int
    version: int
    fields: List[DbfField]


def normalize_key(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    text = text.replace("ة", "ه").replace("ى", "ي").replace("ؤ", "و").replace("ئ", "ي")
    text = text.replace("_", "").replace(" ", "").replace("-", "").replace(".", "")
    return text


def arabic_score(text: str) -> int:
    return len(ARABIC_RE.findall(text or ""))


def read_dbf_meta(path: Path) -> DbfMeta:
    data = path.read_bytes()
    if len(data) < 32:
        raise ValueError(f"ملف DBF تالف: {path}")
    version = data[0]
    record_count = int.from_bytes(data[4:8], "little")
    header_len = int.from_bytes(data[8:10], "little")
    record_len = int.from_bytes(data[10:12], "little")
    ldid = data[29] if len(data) > 29 else 0
    if header_len < 33 or header_len > len(data):
        raise ValueError(f"طول هيدر DBF غير صالح: {header_len}")

    fields: List[DbfField] = []
    offset = 1
    pos = 32
    while pos + 32 <= header_len and data[pos] != 0x0D:
        raw_name = data[pos:pos + 11].split(b"\x00", 1)[0]
        try:
            name = raw_name.decode("ascii")
        except UnicodeDecodeError:
            name = raw_name.decode("latin1")
        name = name.strip()
        ftype = chr(data[pos + 11]).upper()
        flen = data[pos + 16]
        dec = data[pos + 17]
        fields.append(DbfField(name=name, type=ftype, length=flen, decimal=dec, offset=offset))
        offset += flen
        pos += 32

    if not fields:
        raise ValueError("لا توجد حقول في هيدر DBF")
    return DbfMeta(
        path=path,
        header=data[:header_len],
        header_len=header_len,
        record_len=record_len,
        record_count=record_count,
        ldid=ldid,
        version=version,
        fields=fields,
    )


def iter_records(path: Path, meta: DbfMeta) -> Iterable[bytes]:
    data = path.read_bytes()
    start = meta.header_len
    for i in range(meta.record_count):
        a = start + i * meta.record_len
        b = a + meta.record_len
        if b <= len(data):
            yield data[a:b]


def detect_encoding(path: Path, meta: DbfMeta, forced: Optional[str] = None) -> str:
    if forced:
        return forced
    mapped = CODEPAGE_MAP.get(meta.ldid)
    candidates = []
    for enc in (mapped, "cp1256", "cp720", "cp864"):
        if enc and enc not in candidates:
            candidates.append(enc)

    char_fields = [f for f in meta.fields if f.type in ("C", "V")]
    best_enc = candidates[0] if candidates else "cp1256"
    best_score = -1
    samples = list(iter_records(path, meta))[:80]
    if not samples or not char_fields:
        return best_enc

    for enc in candidates:
        score = 0
        for rec in samples:
            if not rec or rec[0:1] == b"*":
                continue
            for field in char_fields:
                chunk = rec[field.offset:field.offset + field.length]
                try:
                    text = chunk.decode(enc, errors="ignore")
                except LookupError:
                    text = ""
                score += arabic_score(text)
        if score > best_score:
            best_score = score
            best_enc = enc
    return best_enc


def cell_to_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "T" if value else "F"
    if isinstance(value, dt.datetime):
        return value.strftime("%Y%m%d")
    if isinstance(value, dt.date):
        return value.strftime("%Y%m%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_date_value(value: Any) -> Optional[dt.date]:
    if value is None or value == "":
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = str(value).strip()
    text = text.replace(".", "-").replace("/", "-").replace("\\", "-")
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y", "%Y%m%d"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    m = re.fullmatch(r"(\d{1,2})-(\d{1,2})-(\d{2,4})", text)
    if m:
        d, mo, y = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if y < 100:
            y += 2000
        try:
            return dt.date(y, mo, d)
        except ValueError:
            return None
    return None


def parse_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace(",", "")
    text = text.replace("٫", ".")
    try:
        return float(text)
    except ValueError:
        return None


def encode_c(value: Any, length: int, encoding: str) -> bytes:
    text = cell_to_text(value)
    raw = text.encode(encoding, errors="replace")
    if len(raw) > length:
        raw = raw[:length]
    return raw.ljust(length, b" ")


def encode_n(value: Any, length: int, decimal: int) -> bytes:
    num = parse_number(value)
    if num is None:
        return b" " * length
    if decimal <= 0:
        s = str(int(round(num)))
    else:
        s = f"{num:.{decimal}f}"
    if len(s) > length:
        return b"*" * length
    return s.encode("ascii").rjust(length, b" ")


def encode_d(value: Any) -> bytes:
    parsed = parse_date_value(value)
    if parsed is None:
        return b" " * 8
    return parsed.strftime("%Y%m%d").encode("ascii")


def encode_l(value: Any) -> bytes:
    text = cell_to_text(value).lower()
    if text in ("t", "true", "1", "y", "yes", "نعم"):
        return b"T"
    if text in ("f", "false", "0", "n", "no", "لا"):
        return b"F"
    return b" "


def encode_field(field: DbfField, value: Any, encoding: str) -> bytes:
    if field.type in ("C", "V", "W"):
        return encode_c(value, field.length, encoding)
    if field.type in ("N", "F", "B", "I", "Y"):
        if field.type == "I" and field.length == 4:
            num = parse_number(value) or 0
            return struct.pack("<i", int(num))
        return encode_n(value, field.length, field.decimal)
    if field.type == "D":
        return encode_d(value)
    if field.type == "L":
        return encode_l(value)
    if field.type == "M":
        return b" " * field.length
    return encode_c(value, field.length, encoding)


def build_record(meta: DbfMeta, values: Dict[str, Any], encoding: str) -> bytes:
    rec = bytearray(b" " * meta.record_len)
    rec[0:1] = b" "
    for field in meta.fields:
        chunk = encode_field(field, values.get(field.name), encoding)
        if len(chunk) != field.length:
            chunk = (chunk + b" " * field.length)[: field.length]
        rec[field.offset:field.offset + field.length] = chunk
    if len(rec) != meta.record_len:
        raise ValueError("طول السجل لا يطابق هيكل DBF")
    return bytes(rec)


def header_with_count(header: bytes, count: int, when: Optional[dt.date] = None) -> bytes:
    out = bytearray(header)
    stamp = when or dt.date.today()
    year = stamp.year - 1900
    if year < 0:
        year = 0
    if year > 255:
        year = year % 100
    out[1] = year
    out[2] = stamp.month
    out[3] = stamp.day
    out[4:8] = struct.pack("<I", count)
    return bytes(out)


def zap_keep_structure(path: Path, meta: DbfMeta) -> None:
    path.write_bytes(header_with_count(meta.header, 0) + b"\x1a")


def write_records(path: Path, meta: DbfMeta, records: Sequence[bytes]) -> None:
    payload = header_with_count(meta.header, len(records))
    payload += b"".join(records)
    payload += b"\x1a"
    path.write_bytes(payload)


def structure_signature(header: bytes) -> bytes:
    """الهيدر بدون تاريخ التحديث وعداد السجلات."""
    sig = bytearray(header)
    sig[1:8] = b"\x00" * 7
    return bytes(sig)


def aliases_for_field(name: str) -> List[str]:
    key = name.upper()
    out = [normalize_key(name)]
    for alias in FIELD_ALIASES.get(key, ()):
        na = normalize_key(alias)
        if na and na not in out:
            out.append(na)
    return out


def match_field(excel_header: Any, fields: Sequence[DbfField]) -> Optional[str]:
    nh = normalize_key(excel_header)
    if not nh:
        return None
    for field in fields:
        if normalize_key(field.name) == nh:
            return field.name
    for field in fields:
        for alias in aliases_for_field(field.name):
            if nh == alias or alias in nh or nh in alias:
                return field.name
    return None


def looks_like_header_row(row: Sequence[Any], fields: Sequence[DbfField]) -> int:
    score = 0
    for cell in row:
        if match_field(cell, fields):
            score += 3
        text = cell_to_text(cell)
        if text and not parse_number(text) and arabic_score(text) + len(text) > 0:
            # صف العناوين غالباً نصوص قصيرة غير أرقام موظفين
            if not DIGITS_RE.fullmatch(text):
                score += 1
    return score


def load_excel_table(excel_path: Path) -> Tuple[List[Any], List[List[Any]], str]:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    while rows and all(c is None or str(c).strip() == "" for c in rows[0]):
        rows.pop(0)
    if not rows:
        raise ValueError("ملف Excel فارغ")
    return rows[0], rows[1:], ws.title


def extract_filename_hints(excel_path: Path) -> Dict[str, Any]:
    name = excel_path.stem
    hints: Dict[str, Any] = {}
    m_date = re.search(r"(\d{1,2})[^\d](\d{1,2})[^\d](\d{2,4})", name)
    if m_date:
        d, mo, y = int(m_date.group(1)), int(m_date.group(2)), int(m_date.group(3))
        if y < 100:
            y += 2000
        try:
            hints["DATE"] = dt.date(y, mo, d)
        except ValueError:
            pass
    nums = DIGITS_RE.findall(name)
    # رقم الأمر غالباً 4 أرقام وليس السنة
    for n in nums:
        if n in {"13", "8", "08", "2026", "26"}:
            continue
        if 3 <= len(n) <= 6:
            hints["AMER"] = n
            break
    return hints


def sanitize_emp_no(value: Any) -> str:
    text = cell_to_text(value)
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    return text.strip()


def pick_key_field(mapping: Dict[int, str]) -> Optional[str]:
    names = list(mapping.values())
    pno_keys = {normalize_key(a) for a in FIELD_ALIASES["PNO"]}
    pno_keys.add("PNO")
    for name in names:
        if normalize_key(name) == "PNO":
            return name
    for name in names:
        if normalize_key(name) in pno_keys:
            return name
    return names[0] if names else None


def row_key_value(row: Sequence[Any], mapping: Dict[int, str], key_field: str) -> str:
    for idx, name in mapping.items():
        if name == key_field:
            return sanitize_emp_no(row[idx] if idx < len(row) else None)
    return ""


def assert_dbf_writable(path: Path) -> None:
    try:
        with path.open("r+b") as f:
            f.seek(0)
            f.read(1)
    except PermissionError as exc:
        raise PermissionError(
            "ملف الفوكس مفتوح. أغلق Visual FoxPro تماماً ثم أعد التشغيل. "
            "لذلك بقي العدد القديم 222."
        ) from exc


def map_and_build_rows(
    excel_path: Path,
    meta: DbfMeta,
    encoding: str,
    expected_count: Optional[int] = None,
) -> Tuple[List[bytes], Dict[str, Any]]:
    first, rest, sheet = load_excel_table(excel_path)
    preview = [first] + rest[:14]
    header_idx = 0
    best = looks_like_header_row(first, meta.fields)
    for i, row in enumerate(preview):
        score = looks_like_header_row(row, meta.fields)
        if score > best:
            best = score
            header_idx = i
    all_rows = [first] + rest
    headers = all_rows[header_idx]
    data_rows = all_rows[header_idx + 1:]

    mapping: Dict[int, str] = {}
    used = set()
    for i, header in enumerate(headers):
        field_name = match_field(header, meta.fields)
        if field_name and field_name not in used:
            mapping[i] = field_name
            used.add(field_name)

    positional = False
    excel_width = 0
    if headers:
        excel_width = max((i + 1 for i, h in enumerate(headers) if h not in (None, "")), default=0)
    if not mapping and excel_width == len(meta.fields):
        positional = True
        mapping = {i: meta.fields[i].name for i in range(len(meta.fields))}
        used = {f.name for f in meta.fields}

    if not mapping:
        raise ValueError(
            "تعذر مطابقة أعمدة Excel مع حقول DBF. "
            f"عناوين Excel: {headers} | حقول الجدول: {[f.name for f in meta.fields]}"
        )

    hints = extract_filename_hints(excel_path)
    hint_values: Dict[str, Any] = {}
    for field in meta.fields:
        if field.name in used:
            continue
        key = field.name.upper()
        if key in hints:
            hint_values[field.name] = hints[key]
        elif key in ("DATE", "DT", "FDATE") and "DATE" in hints:
            hint_values[field.name] = hints["DATE"]
        elif key in ("AMER", "AMR", "ORDER", "ORDNO", "NOAMR") and "AMER" in hints:
            hint_values[field.name] = hints["AMER"]

    key_field = pick_key_field(mapping)
    records: List[bytes] = []
    skipped = 0
    started = False
    blank_run = 0
    extra_after_expected = 0

    for row in data_rows:
        if row is None:
            skipped += 1
            blank_run += 1
            if started and blank_run >= 2:
                break
            continue
        if all(c is None or str(c).strip() == "" for c in row):
            skipped += 1
            blank_run += 1
            if started and blank_run >= 2:
                break
            continue

        key_val = row_key_value(row, mapping, key_field) if key_field else ""
        if key_field and not key_val:
            skipped += 1
            blank_run += 1
            if started and blank_run >= 3:
                break
            continue

        blank_run = 0
        started = True
        if expected_count is not None and len(records) >= expected_count:
            extra_after_expected += 1
            continue

        values: Dict[str, Any] = dict(hint_values)
        for idx, field_name in mapping.items():
            values[field_name] = row[idx] if idx < len(row) else None
        if key_field and not cell_to_text(values.get(key_field)):
            skipped += 1
            continue
        records.append(build_record(meta, values, encoding))

    if expected_count is not None:
        if len(records) < expected_count:
            raise ValueError(
                f"عدد القيود المستوردة {len(records)} أقل من المطلوب {expected_count}"
            )
        if extra_after_expected:
            print("تم تجاهل صفوف إضافية بعد العدد المطلوب:", extra_after_expected)

    report = {
        "sheet": sheet,
        "headers": headers,
        "mapping": {headers[i] if i < len(headers) else i: name for i, name in mapping.items()},
        "positional": positional,
        "hints": {k: cell_to_text(v) for k, v in hint_values.items()},
        "inserted": len(records),
        "skipped": skipped,
        "header_row": header_idx + 1,
        "key_field": key_field,
        "excel_data_rows": len(data_rows),
        "extra_after_expected": extra_after_expected,
        "expected_count": expected_count,
    }
    return records, report


def find_named_file(roots: Sequence[Path], wanted: str) -> Optional[Path]:
    wanted_u = wanted.upper()
    for root in roots:
        if not root.exists():
            continue
        if root.is_file() and root.name.upper() == wanted_u:
            return root
        try:
            for p in root.rglob("*"):
                if p.is_file() and p.name.upper() == wanted_u:
                    return p
        except OSError:
            continue
    return None


def find_excel(roots: Sequence[Path]) -> Optional[Path]:
    preferred = []
    others = []
    for root in roots:
        if not root.exists():
            continue
        try:
            iterator = [root] if root.is_file() else root.rglob("*")
            for p in iterator:
                if not p.is_file():
                    continue
                if p.suffix.lower() not in (".xlsx", ".xlsm"):
                    continue
                name = p.name
                if "علاوات" in name or "8276" in name or "alaw" in name.lower():
                    preferred.append(p)
                else:
                    others.append(p)
        except OSError:
            continue
    if preferred:
        return preferred[0]
    return others[0] if others else None


def default_search_roots() -> List[Path]:
    here = Path(__file__).resolve().parent
    desktop = Path.home() / "Desktop"
    roots = [
        Path(r"C:\Users\ngc\Desktop\092026\NewٌRel"),
        Path(r"C:\Users\ngc\Desktop\092026"),
        desktop / "092026" / "NewٌRel",
        desktop / "092026",
        here / "data",
        here,
    ]
    # إن وُجد مجلد 092026 بأي تهجئة
    if desktop.exists():
        for p in desktop.iterdir():
            if p.is_dir() and "092026" in p.name:
                roots.append(p)
                roots.extend([c for c in p.iterdir() if c.is_dir()])
    return roots


def resolve_paths(excel_arg: Optional[str], dbf_arg: Optional[str]) -> Tuple[Path, Path]:
    roots = default_search_roots()
    excel: Optional[Path] = Path(excel_arg) if excel_arg else None
    dbfp: Optional[Path] = Path(dbf_arg) if dbf_arg else None

    if excel and not excel.exists():
        excel = None
    if dbfp and not dbfp.exists():
        dbfp = None

    if dbfp is None:
        dbfp = find_named_file(roots, "AL082026.DBF")
    if excel is None:
        # مسار المستخدم الحرفي أولاً
        exact = Path(r"C:\Users\ngc\Desktop\092026\NewٌRel") / "علاوات حسب امر   8276في 13-8-2026.xlsx"
        excel = exact if exact.exists() else find_excel(roots)

    if excel is None or dbfp is None:
        raise FileNotFoundError(
            "تعذر إيجاد الملفات. ضع Excel و AL082026.DBF في مجلد data "
            "أو مرّر --excel و --dbf"
        )
    return excel, dbfp


def inspect(excel_path: Path, dbf_path: Path) -> None:
    print("=" * 60)
    print("فحص الملفات")
    print("=" * 60)
    print("Excel:")
    print(str(excel_path))
    print("موجود:" , "نعم" if excel_path.exists() else "لا")
    if excel_path.exists():
        first, rest, sheet = load_excel_table(excel_path)
        print("الورقة:", sheet)
        print("عدد الصفوف:", 1 + len(rest))
        print("صف 1:", first)
        if rest:
            print("صف 2:", rest[0])

    print()
    print("DBF:")
    print(str(dbf_path))
    print("موجود:" , "نعم" if dbf_path.exists() else "لا")
    if dbf_path.exists():
        meta = read_dbf_meta(dbf_path)
        enc = detect_encoding(dbf_path, meta)
        print("Language Driver:", meta.ldid, f"(0x{meta.ldid:02X})")
        print("الترميز:", enc)
        print("عدد السجلات الحالي:", meta.record_count)
        print("طول السجل:", meta.record_len)
        print("الحقول:")
        for f in meta.fields:
            print(f"  {f.name:10} {f.type}({f.length},{f.decimal})")
        if excel_path.exists():
            try:
                _recs, report = map_and_build_rows(excel_path, meta, enc, expected_count=None)
                print("قيود Excel الصالحة للنقل:", report["inserted"])
                print("سجلات الفوكس الحالية:", meta.record_count)
                if report["inserted"] != meta.record_count:
                    print("تنبيه: العددان غير متطابقين. بعد النقل يجب أن يصبح الفوكس:", report["inserted"])
            except Exception as exc:
                print("تعذر عد قيود Excel:", exc)
        for rec in list(iter_records(dbf_path, meta))[:2]:
            sample = {}
            for f in meta.fields:
                if f.type in ("C", "V"):
                    sample[f.name] = rec[f.offset:f.offset + f.length].decode(enc, errors="replace").rstrip()
                else:
                    sample[f.name] = rec[f.offset:f.offset + f.length].decode("latin1", errors="replace").strip()
            print("سجل نموذجي:", sample)


def backup_files(dbf_path: Path) -> Path:
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = dbf_path.with_name(dbf_path.name + f".bak_{stamp}")
    shutil.copy2(dbf_path, backup)
    for ext in (".FPT", ".fpt", ".DBT", ".dbt", ".CDX", ".cdx", ".IDX", ".idx"):
        extra = dbf_path.with_suffix(ext)
        if extra.exists():
            shutil.copy2(extra, Path(str(backup) + extra.suffix))
    return backup


def transfer(
    excel_path: Path,
    dbf_path: Path,
    encoding_forced: Optional[str] = None,
    dry_run: bool = False,
    expected_count: Optional[int] = None,
) -> Dict[str, Any]:
    if not excel_path.exists():
        raise FileNotFoundError(f"ملف Excel غير موجود: {excel_path}")
    if not dbf_path.exists():
        raise FileNotFoundError(f"ملف DBF غير موجود: {dbf_path}")

    meta = read_dbf_meta(dbf_path)
    encoding = detect_encoding(dbf_path, meta, encoding_forced)
    original_sig = structure_signature(meta.header)
    records, report = map_and_build_rows(
        excel_path, meta, encoding, expected_count=expected_count
    )
    if not records:
        raise ValueError("لا توجد صفوف قابلة للنقل من Excel")

    print("=" * 60)
    print("خطة النقل")
    print("=" * 60)
    print("من:")
    print(str(excel_path))
    print("إلى:")
    print(str(dbf_path))
    print("ترميز العربي:", encoding)
    print("Language Driver:", f"0x{meta.ldid:02X}")
    print("صف العناوين:", report["header_row"])
    print("مطابقة الأعمدة:", report["mapping"])
    print("مطابقة موضعية:", "نعم" if report["positional"] else "لا")
    if report["hints"]:
        print("قيم من اسم الملف:", report["hints"])
    print("سجلات قديمة في الفوكس:", meta.record_count)
    print("قيود Excel بعد العنوان:", report["inserted"])
    print("صفوف متجاوزة:", report["skipped"])
    print("حقل المفتاح:", report.get("key_field"))
    if expected_count is not None:
        print("العدد المطلوب:", expected_count)

    if dry_run:
        print("وضع التجربة: لم يُكتب شيء.")
        return report

    locked = False
    try:
        assert_dbf_writable(dbf_path)
    except PermissionError:
        locked = True

    if locked:
        sidecar = dbf_path.with_name("AL082026_NEW.DBF")
        print("الملف مفتوح في Visual FoxPro لذلك لا يمكن تفريغ 222 من بايثون.")
        print("تم إنشاء ملف جديد بنفس الهيكل:")
        print(str(sidecar))
        write_records(sidecar, meta, records)
        report["backup"] = ""
        report["encoding"] = encoding
        report["old_count"] = meta.record_count
        report["new_count"] = len(records)
        report["sidecar"] = str(sidecar)
        print("من نافذة Command داخل نفس فوكس برو نفّذ:")
        print("DO transfer_alawat.prg")
        print("أو:")
        print("SELECT AL082026")
        print("USE")
        print("USE AL082026 EXCLUSIVE")
        print("ZAP")
        print("APPEND FROM AL082026_NEW")
        print("COUNT")
        print("يجب أن يظهر 126")
        return report

    backup = backup_files(dbf_path)
    print("نسخة احتياطية:")
    print(str(backup))

    zap_keep_structure(dbf_path, meta)
    emptied = read_dbf_meta(dbf_path)
    if emptied.record_count != 0:
        raise RuntimeError("فشل تفريغ الجدول")
    if structure_signature(emptied.header) != original_sig:
        raise RuntimeError("تغيرت هيكلية الجدول أثناء التفريغ")

    write_records(dbf_path, meta, records)
    final = read_dbf_meta(dbf_path)
    if structure_signature(final.header) != original_sig:
        raise RuntimeError("تغيرت هيكلية الجدول بعد الكتابة")
    if final.record_count != len(records):
        raise RuntimeError("عدد السجلات بعد النقل غير مطابق")
    if final.ldid != meta.ldid:
        # أعد بايت اللغة كما كان
        with dbf_path.open("r+b") as f:
            f.seek(29)
            f.write(bytes([meta.ldid]))

    report["backup"] = str(backup)
    report["encoding"] = encoding
    report["old_count"] = meta.record_count
    report["new_count"] = final.record_count
    print("-" * 60)
    print("تم التفريغ ثم النقل.")
    print("عدد الفوكس قبل:", meta.record_count)
    print("عدد الفوكس بعد:", final.record_count)
    print("عدد قيود Excel المنقولة:", report["inserted"])
    if final.record_count != report["inserted"]:
        raise RuntimeError("عدد الفوكس لا يساوي عدد قيود Excel")
    if expected_count is not None and final.record_count != expected_count:
        raise RuntimeError(
            f"العدد بعد النقل {final.record_count} وليس {expected_count}"
        )
    print("الهيكل كما هو، والترميز كما في الجدول الأصلي.")
    return report


def write_report(excel_path: Path, dbf_path: Path, report: Dict[str, Any]) -> Path:
    out = dbf_path.with_name("AL082026_transfer_report.txt")
    lines = [
        "تقرير نقل العلاوات من Excel إلى AL082026.DBF",
        "=" * 60,
        f"Excel: {excel_path}",
        f"DBF: {dbf_path}",
        f"الترميز: {report.get('encoding', '')}",
        f"سجلات قديمة في الفوكس: {report.get('old_count', '')}",
        f"سجلات جديدة في الفوكس: {report.get('new_count', report.get('inserted', ''))}",
        f"العدد المطلوب: {report.get('expected_count', '')}",
        f"صفوف متجاوزة: {report.get('skipped', '')}",
        f"حقل المفتاح: {report.get('key_field', '')}",
        f"صف العناوين: {report.get('header_row', '')}",
        f"مطابقة الأعمدة: {report.get('mapping', '')}",
        f"قيم إضافية: {report.get('hints', '')}",
        f"نسخة احتياطية: {report.get('backup', '')}",
        "",
    ]
    out.write_text("\n".join(lines), encoding="utf-8")
    return out


def make_test_dbf(path: Path, records: Sequence[Tuple[str, str, float]], ldid: int = 0x7E) -> None:
    """DBF داتابيس III بحقول ثابتة لاختبار النقل."""
    fields = [
        ("PNO", "C", 6, 0),
        ("NAME", "C", 36, 0),
        ("AMT", "N", 12, 3),
        ("AMER", "C", 10, 0),
        ("DT", "D", 8, 0),
    ]
    header_len = 32 + 32 * len(fields) + 1
    rec_len = 1 + sum(f[2] for f in fields)
    header = bytearray(header_len)
    header[0] = 0x03
    header[1:4] = bytes([126, 8, 13])
    header[4:8] = struct.pack("<I", 0)
    header[8:10] = struct.pack("<H", header_len)
    header[10:12] = struct.pack("<H", rec_len)
    header[29] = ldid
    pos = 32
    for name, ftype, length, dec in fields:
        raw = name.encode("ascii").ljust(11, b"\x00")
        header[pos:pos + 11] = raw[:11]
        header[pos + 11] = ord(ftype)
        header[pos + 16] = length
        header[pos + 17] = dec
        pos += 32
    header[pos] = 0x0D
    meta = read_dbf_meta_from_bytes(bytes(header), path)
    built = []
    for pno, name, amt in records:
        built.append(
            build_record(
                meta,
                {"PNO": pno, "NAME": name, "AMT": amt, "AMER": "1111", "DT": dt.date(2026, 1, 1)},
                "cp1256",
            )
        )
    write_records(path, meta, built)


def read_dbf_meta_from_bytes(header: bytes, path: Path) -> DbfMeta:
    tmp = path
    tmp.write_bytes(header + b"\x1a")
    return read_dbf_meta(tmp)


def make_test_excel(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "علاوات حسب امر 8276 في 13-8-2026"
    ws.append(["رقم الموظف", "الاسم", "المبلغ"])
    ws.append(["100045", "ابراهيم حمود عبد محمد", 125000])
    ws.append(["110176", "ابراهيم محمد عابد عبود", 98000.5])
    ws.append([None, None, None])
    ws.append([None, None, None])
    for i in range(96):
        ws.append([str(800000 + i), "قيد قديم فائض", 1])
    wb.save(path)


def self_test(base: Optional[Path] = None) -> int:
    import tempfile

    if base is None:
        tmp = tempfile.TemporaryDirectory(prefix="al082026_")
        base = Path(tmp.name)
    else:
        tmp = None
        base.mkdir(parents=True, exist_ok=True)
    dbf_path = base / "AL082026.DBF"
    excel_path = base / "علاوات حسب امر   8276في 13-8-2026.xlsx"
    make_test_dbf(
        dbf_path,
        [("999999", "سجل قديم يجب حذفه", 1), ("888888", "سجل قديم آخر", 2)],
        ldid=0x7E,
    )
    before = read_dbf_meta(dbf_path)
    old_sig = structure_signature(before.header)
    make_test_excel(excel_path)
    report = transfer(excel_path, dbf_path)
    after = read_dbf_meta(dbf_path)
    assert structure_signature(after.header) == old_sig, "تغير الهيكل"
    assert after.ldid == 0x7E, "تغير Language Driver"
    assert after.record_count == 2, after.record_count
    recs = list(iter_records(dbf_path, after))
    names = []
    pnos = []
    for rec in recs:
        pnos.append(rec[1:7].decode("cp1256").strip())
        names.append(rec[7:43].decode("cp1256").rstrip())
        amer = rec[55:65].decode("cp1256").strip()
        dateb = rec[65:73].decode("ascii")
        assert amer == "8276"
        assert dateb == "20260813"
    assert "999999" not in pnos
    assert "100045" in pnos
    assert "ابراهيم حمود عبد محمد" in names
    print("SELF-TEST OK", report["inserted"], names)

    # صفوف قديمة متصلة بدون فراغ: لا تُؤخذ أكثر من العدد المطلوب
    excel2 = base / "extra.xlsx"
    dbf2 = base / "AL2.DBF"
    make_test_dbf(dbf2, [("999999", "قديم", 1)] * 222, ldid=0)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["رقم الموظف", "الاسم", "المبلغ"])
    ws.append(["100045", "ابراهيم حمود عبد محمد", 125000])
    ws.append(["110176", "ابراهيم محمد عابد عبود", 98000.5])
    for i in range(220):
        ws.append([str(700000 + i), "فائض", 1])
    wb.save(excel2)
    report2 = transfer(excel2, dbf2, expected_count=2)
    after2 = read_dbf_meta(dbf2)
    assert after2.record_count == 2, after2.record_count
    assert report2["extra_after_expected"] == 220
    print("COUNT-LIMIT OK", after2.record_count, "from old", 222)

    dbf3 = base / "LOCK.DBF"
    make_test_dbf(dbf3, [("999999", "قديم", 1)], ldid=0)
    os.chmod(dbf3, 0o444)
    try:
        report3 = transfer(excel_path, dbf3, expected_count=2)
        sidecar = dbf3.with_name("AL082026_NEW.DBF")
        assert sidecar.exists(), "sidecar missing"
        meta_side = read_dbf_meta(sidecar)
        assert meta_side.record_count == 2, meta_side.record_count
        assert read_dbf_meta(dbf3).record_count == 1
        print("SIDECAR-WHEN-LOCKED OK", meta_side.record_count)
    finally:
        os.chmod(dbf3, 0o644)

    if tmp is not None:
        tmp.cleanup()
    return 0


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="تفريغ AL082026.DBF ونقل بيانات Excel إليه")
    p.add_argument("--excel", default=None, help="مسار ملف Excel")
    p.add_argument("--dbf", default=None, help="مسار AL082026.DBF")
    p.add_argument("--encoding", default=None, help="فرض ترميز مثل cp1256")
    p.add_argument("--expected-count", type=int, default=None, help="عدد القيود المطلوب مثل 126")
    p.add_argument("--inspect-only", action="store_true", help="عرض البنية فقط")
    p.add_argument("--dry-run", action="store_true", help="تجربة بدون كتابة")
    p.add_argument("--self-test", action="store_true", help="اختبار ذاتي بملفات مؤقتة")
    return p.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    if args.self_test:
        return self_test(None)

    try:
        excel_path, dbf_path = resolve_paths(args.excel, args.dbf)
    except Exception as exc:
        print("خطأ:", exc)
        return 1

    if os.name == "nt":
        excel_path = Path(os.path.expandvars(str(excel_path)))
        dbf_path = Path(os.path.expandvars(str(dbf_path)))

    if args.inspect_only:
        inspect(excel_path, dbf_path)
        return 0

    try:
        report = transfer(
            excel_path=excel_path,
            dbf_path=dbf_path,
            encoding_forced=args.encoding,
            dry_run=args.dry_run,
            expected_count=args.expected_count,
        )
        if not args.dry_run:
            out = write_report(excel_path, dbf_path, report)
            print("التقرير:")
            print(str(out))
    except Exception as exc:
        print("\nخطأ:", exc)
        print("أغلق نظام الأفراد/الفوكس إن كان الملف مفتوحاً ثم أعد التشغيل.")
        print("أو نفّذ الفحص:")
        print('  python transfer_excel_to_al_dbf.py --inspect-only')
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
