# -*- coding: utf-8 -*-
"""نقل مدة الخدمة (DD/MM/YY) من ser.xls إلى MAIN720.xls حسب الاسم."""

from pathlib import Path
import re
import xlrd
from xlutils.copy import copy as xl_copy


def norm(s):
    t = str(s or "").strip()
    t = t.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    t = t.replace("ة", "ه").replace("ى", "ي").replace("عبد ال", "عبدال")
    return re.sub(r"\s+", " ", t)


def score_match(ser_name, main_name):
    sn, mn = norm(ser_name), norm(main_name)
    if not sn or not mn:
        return 0
    if sn == mn:
        return 100
    if mn.startswith(sn) or sn.startswith(mn):
        return 90
    st, mt = sn.split(), mn.split()
    if len(st) >= 3 and len(mt) >= 3 and st[:3] == mt[:3]:
        return 85
    if len(st) >= 2 and len(mt) >= 2 and st[:2] == mt[:2]:
        return 70
    st2 = [x.replace("شكور", "شكر").replace("حسيب", "حسين").replace("كريم", "كرم").replace("حمه", "حمد") for x in st]
    mt2 = [x.replace("شكور", "شكر").replace("حسيب", "حسين").replace("كريم", "كرم").replace("حمه", "حمد") for x in mt]
    if len(st2) >= 2 and st2[:2] == mt2[:2]:
        return 75
    if len(st) >= 4 and len(mt) >= 4 and st[0] == mt[0] and st[2] == mt[2] and st[3] == mt[3]:
        return 72
    return 0


def transfer(main_path: Path, ser_path: Path, out_path: Path):
    main_book = xlrd.open_workbook(str(main_path), formatting_info=True)
    ser_book = xlrd.open_workbook(str(ser_path))
    msh, ssh = main_book.sheet_by_index(0), ser_book.sheet_by_index(0)
    headers = [str(x).strip().upper() for x in msh.row_values(0)]
    name_i, dd_i, mm_i, yy_i = map(headers.index, ("NAME", "DD", "MM", "YY"))
    mains = [{"row": r, "name": msh.cell_value(r, name_i)} for r in range(1, msh.nrows)]
    wb = xl_copy(main_book)
    ws = wb.get_sheet(0)
    updated = 0
    for r in range(1, ssh.nrows):
        sname, sdd, smm, syy = ssh.row_values(r)[:4]
        scored = sorted(((score_match(sname, m["name"]), m) for m in mains), key=lambda x: x[0], reverse=True)
        if not scored or scored[0][0] < 65:
            print("UNMATCHED:", sname)
            continue
        best = scored[0][1]
        for col, val in ((dd_i, sdd), (mm_i, smm), (yy_i, syy)):
            ws.write(best["row"], col, float(val) if val != "" else "")
        updated += 1
        print(f"OK {sname} => {best['name']} ({sdd}/{smm}/{syy})")
    wb.save(str(out_path))
    print("updated=", updated, "out=", out_path)


if __name__ == "__main__":
    transfer(
        Path("data/MAIN720_original.xls"),
        Path("data/ser.xls"),
        Path("data/MAIN720_updated.xls"),
    )
