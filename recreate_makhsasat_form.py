# -*- coding: utf-8 -*-
"""
إنشاء استمارة مخصصات جديدة بنفس قياس ومواصفات الملف الأصلي إن وُجد.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Optional, Tuple

from reportlab.lib.colors import Color, black, white
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
except ImportError:
    print("ثبّت: pip install arabic-reshaper python-bidi reportlab pymupdf")
    sys.exit(1)


NAVY = Color(0.10, 0.18, 0.32)
LINE = Color(0.15, 0.15, 0.15)
PALE = Color(0.93, 0.94, 0.96)
RULE = Color(0.55, 0.42, 0.18)

ORIGINAL_CANDIDATES = [
    Path(r"C:\Users\ngc\Desktop\092026\مخصصات.pdf"),
    Path.home() / "Desktop" / "092026" / "مخصصات.pdf",
    Path(__file__).resolve().parent / "data" / "مخصصات.pdf",
    Path(__file__).resolve().parent / "مخصصات.pdf",
]


def ar(text: str) -> str:
    return get_display(arabic_reshaper.reshape(text))


def find_font() -> str:
    candidates = [
        Path("/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf"),
        Path("/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf"),
        Path(r"C:\Windows\Fonts\arial.ttf"),
        Path(r"C:\Windows\Fonts\tahoma.ttf"),
        Path(r"C:\Windows\Fonts\trado.ttf"),
        Path(r"C:\Windows\Fonts\arialuni.ttf"),
    ]
    windir = os.environ.get("WINDIR")
    if windir:
        fonts = Path(windir) / "Fonts"
        candidates.extend(
            [
                fonts / "arial.ttf",
                fonts / "tahoma.ttf",
                fonts / "trado.ttf",
                fonts / "NotoNaskhArabic-Regular.ttf",
            ]
        )
    for p in candidates:
        if p.exists():
            pdfmetrics.registerFont(TTFont("Arab", str(p)))
            bold = p.with_name(p.name.replace("Regular", "Bold"))
            if bold.exists():
                pdfmetrics.registerFont(TTFont("ArabBold", str(bold)))
            else:
                pdfmetrics.registerFont(TTFont("ArabBold", str(p)))
            return str(p)
    raise FileNotFoundError("لم يُعثر على خط عربي")


def find_original() -> Optional[Path]:
    for p in ORIGINAL_CANDIDATES:
        if p.exists():
            return p
    desktop = Path.home() / "Desktop" / "092026"
    if desktop.exists():
        for p in desktop.glob("*.pdf"):
            name = p.name
            if "جديد" in name:
                continue
            if name == "مخصصات.pdf" or "مخصصات.pdf" in name:
                return p
    return None


def read_page_size(pdf_path: Path) -> Tuple[float, float]:
    try:
        import fitz
    except ImportError:
        from pypdf import PdfReader

        reader = PdfReader(str(pdf_path))
        box = reader.pages[0].mediabox
        return float(box.width), float(box.height)
    doc = fitz.open(str(pdf_path))
    rect = doc[0].rect
    size = (float(rect.width), float(rect.height))
    doc.close()
    return size


def box(c: canvas.Canvas, x, y, w, h, fill=None, stroke=LINE, width=0.8):
    c.saveState()
    if fill:
        c.setFillColor(fill)
        c.rect(x, y, w, h, fill=1, stroke=0)
    c.setStrokeColor(stroke)
    c.setLineWidth(width)
    c.rect(x, y, w, h, fill=0, stroke=1)
    c.restoreState()


def hline(c, x1, x2, y, width=0.6, color=LINE):
    c.setStrokeColor(color)
    c.setLineWidth(width)
    c.line(x1, y, x2, y)


def right(c, text, x, y, size=11, font="Arab", color=black):
    c.setFillColor(color)
    c.setFont(font, size)
    c.drawRightString(x, y, ar(text))


def center(c, text, x, y, size=12, font="ArabBold", color=black):
    c.setFillColor(color)
    c.setFont(font, size)
    c.drawCentredString(x, y, ar(text))


def dotted_field(c, x_left, x_right, y):
    c.setStrokeColor(Color(0.55, 0.55, 0.55))
    c.setDash(1, 2)
    c.setLineWidth(0.5)
    c.line(x_left, y - 2, x_right, y - 2)
    c.setDash()


def draw_form(c: canvas.Canvas, width: float, height: float) -> None:
    margin = 28
    inner = 6
    left = margin
    right_x = width - margin
    top = height - margin
    bottom = margin

    box(c, left, bottom, right_x - left, top - bottom, stroke=NAVY, width=1.8)
    box(c, left + inner, bottom + inner, right_x - left - 2 * inner, top - bottom - 2 * inner, stroke=NAVY, width=0.7)

    cx = width / 2.0
    y = top - 22
    center(c, "جمهورية العراق", cx, y, 16, "ArabBold", NAVY)
    y -= 16
    c.setStrokeColor(RULE)
    c.setLineWidth(1.4)
    c.line(cx - 70, y + 6, cx + 70, y + 6)
    y -= 6
    center(c, "استمارة مخصصات", cx, y, 18, "ArabBold", NAVY)
    y -= 15
    center(c, "لصرف المخصصات والعلاوات حسب الأوامر الإدارية", cx, y, 10, "Arab", Color(0.25, 0.25, 0.25))

    y -= 14
    hline(c, left + 14, right_x - 14, y, 1.0, RULE)
    y -= 22

    # meta row
    right(c, "اسم الدائرة :", right_x - 18, y, 11, "ArabBold")
    dotted_field(c, left + 18, right_x - 95, y)
    y -= 20
    right(c, "القسم / الشعبة :", right_x - 18, y, 11, "ArabBold")
    dotted_field(c, left + 18, right_x - 110, y)
    y -= 20
    right(c, "الشهر :", right_x - 18, y, 11, "ArabBold")
    dotted_field(c, right_x - 170, right_x - 70, y)
    right(c, "السنة :", right_x - 190, y, 11, "ArabBold")
    dotted_field(c, left + 18, right_x - 240, y)

    y -= 18
    hline(c, left + 14, right_x - 14, y, 0.6, LINE)
    y -= 18
    right(c, "أولاً : معلومات الموظف", right_x - 18, y, 12, "ArabBold", NAVY)
    y -= 20

    pairs = [
        ("اسم الموظف الرباعي :", "رقم الموظف :"),
        ("العنوان الوظيفي :", "الدرجة / المرحلة :"),
        ("الحالة الاجتماعية :", "عدد الأولاد المستحقين :"),
    ]
    for lab_r, lab_l in pairs:
        mid = left + (right_x - left) * 0.50
        c.setFont("Arab", 11)
        wr = c.stringWidth(ar(lab_r), "Arab", 11) + 8
        wl = c.stringWidth(ar(lab_l), "Arab", 11) + 8
        right(c, lab_r, right_x - 18, y, 11)
        dotted_field(c, mid + 10, right_x - 18 - wr, y)
        right(c, lab_l, mid - 6, y, 11)
        dotted_field(c, left + 18, mid - 6 - wl, y)
        y -= 20

    y -= 4
    hline(c, left + 14, right_x - 14, y, 0.6, LINE)
    y -= 18
    right(c, "ثانياً : تفاصيل المخصصات", right_x - 18, y, 12, "ArabBold", NAVY)
    y -= 10

    table_top = y
    table_left = left + 16
    table_right = right_x - 16
    row_h = 22
    header_h = 24
    rows = 8
    cols = [
        ("ملاحظات", 0),
        ("المبلغ / دينار", 90),
        ("تاريخ الأمر", 78),
        ("رقم الأمر", 78),
        ("نوع المخصص / العلاوة", 170),
        ("ت", 28),
    ]
    total_fixed = sum(w for _, w in cols if w)
    remain = (table_right - table_left) - total_fixed
    widths = [w if w else remain for _, w in cols]

    # header
    x = table_left
    box(c, table_left, table_top - header_h - rows * row_h, table_right - table_left, header_h + rows * row_h, fill=white, width=1.0)
    box(c, table_left, table_top - header_h, table_right - table_left, header_h, fill=PALE, width=1.0)
    x = table_left
    for (title, _), w in zip(cols, widths):
        x2 = x + w
        c.setStrokeColor(LINE)
        c.setLineWidth(0.6)
        c.line(x2, table_top - header_h - rows * row_h, x2, table_top)
        center(c, title, (x + x2) / 2, table_top - 16, 9, "ArabBold", NAVY)
        x = x2
    # numbers and grid
    for r in range(rows):
        yrow = table_top - header_h - (r + 1) * row_h
        hline(c, table_left, table_right, yrow, 0.4, Color(0.7, 0.7, 0.7))
        center(c, str(r + 1), table_right - widths[-1] / 2, yrow + 7, 9, "Arab", Color(0.3, 0.3, 0.3))

    y = table_top - header_h - rows * row_h - 22
    c.setFont("ArabBold", 11)
    s_sum = ar("المجموع رقماً :")
    s_wrd = ar("كتابةً :")
    w_sum = c.stringWidth(s_sum, "ArabBold", 11)
    w_wrd = c.stringWidth(s_wrd, "ArabBold", 11)
    right(c, "المجموع رقماً :", right_x - 18, y, 11, "ArabBold")
    box_w = 92
    box_x = right_x - 22 - w_sum - box_w
    box(c, box_x, y - 6, box_w, 18, width=0.7)
    right(c, "كتابةً :", box_x - 12, y, 11, "ArabBold")
    dotted_field(c, left + 18, box_x - 16 - w_wrd, y)

    y -= 24
    right(c, "رقم الأمر الإداري المعتمد :", right_x - 18, y, 11, "ArabBold")
    dotted_field(c, right_x - 280, right_x - 165, y)
    right(c, "تاريخ الأمر :", right_x - 300, y, 11, "ArabBold")
    dotted_field(c, left + 18, right_x - 380, y)

    y -= 18
    hline(c, left + 14, right_x - 14, y, 0.6, LINE)
    y -= 16
    right(c, "ثالثاً : إقرار وتعهد", right_x - 18, y, 12, "ArabBold", NAVY)
    y -= 18
    pledge = (
        "أقر بأن المعلومات المدونة أعلاه صحيحة ومطابقة للواقع، وأتحمل المسؤولية القانونية "
        "عن أي بيان غير صحيح، وأتعهد بإبلاغ الدائرة عن أي تغيير يطرأ خلال ثلاثين يوماً."
    )
    c.setFont("Arab", 10)
    # wrap RTL by splitting
    words = pledge.split(" ")
    lines = []
    cur = ""
    max_w = table_right - table_left - 8
    for w in words:
        trial = (cur + " " + w).strip()
        if c.stringWidth(ar(trial), "Arab", 10) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    for line in lines:
        right(c, line, right_x - 18, y, 10, "Arab")
        y -= 14

    y -= 8
    hline(c, left + 14, right_x - 14, y, 0.6, LINE)
    y -= 18
    right(c, "رابعاً : التواقيع والاعتماد", right_x - 18, y, 12, "ArabBold", NAVY)
    y -= 26

    sig_titles = [
        "المدير المخول",
        "الحسابات",
        "مسؤول الذاتية / الأفراد",
        "الموظف",
    ]
    sig_w = (table_right - table_left - 18) / 4.0
    sig_h = 78
    sig_y = y - sig_h
    if sig_y < bottom + 36:
        sig_y = bottom + 36
        sig_h = y - sig_y
    for i, title in enumerate(sig_titles):
        x = table_left + i * (sig_w + 6)
        box(c, x, sig_y, sig_w, sig_h, fill=white, width=0.8)
        center(c, title, x + sig_w / 2, sig_y + sig_h - 14, 9, "ArabBold", NAVY)
        right(c, "الاسم :", x + sig_w - 8, sig_y + 42, 8)
        dotted_field(c, x + 8, x + sig_w - 40, sig_y + 42)
        right(c, "التوقيع :", x + sig_w - 8, sig_y + 24, 8)
        dotted_field(c, x + 8, x + sig_w - 48, sig_y + 24)
        right(c, "التاريخ :", x + sig_w - 8, sig_y + 8, 8)
        dotted_field(c, x + 8, x + sig_w - 48, sig_y + 8)

    center(
        c,
        "نسخة إلى : الحسابات  —  إضبارة الموظف  —  الذاتية",
        cx,
        bottom + 14,
        8,
        "Arab",
        Color(0.3, 0.3, 0.3),
    )


def draw_notes_page(c: canvas.Canvas, width: float, height: float) -> None:
    margin = 28
    left = margin
    right_x = width - margin
    top = height - margin
    bottom = margin
    box(c, left, bottom, right_x - left, top - bottom, stroke=NAVY, width=1.8)
    cx = width / 2.0
    y = top - 28
    center(c, "ملاحظات وتعليمات استمارة المخصصات", cx, y, 14, "ArabBold", NAVY)
    y -= 16
    hline(c, left + 24, right_x - 24, y, 1.0, RULE)
    y -= 24
    notes = [
        "تعتمد البيانات المدونة في هذه الاستمارة أساساً لاستحقاق وصرف المخصصات وفق الأوامر الإدارية النافذة.",
        "لا تُصرف المخصصات إلا بعد استكمال المستمسكات المؤيدة ومطابقة رقم الموظف مع سجلات النظام.",
        "يُذكر رقم الأمر الإداري وتاريخه إزاء كل مخصص، ويُمنع تكرار الصرف عن نفس الأمر.",
        "المجموع رقماً وكتابةً يجب أن يطابق مجموع حقول المبالغ في الجدول.",
        "أي تعديل على الاستمارة بعد اعتمادها يتطلب تأشيرة المخول بالحسابات.",
        "تُحفظ نسخة في إضبارة الموظف ونسخة لدى الحسابات.",
        "الاستمارة مطابقة لقياس الملف الأصلي عند وجوده، ومعدة للطباعة على نفس القياس.",
    ]
    for i, note in enumerate(notes, 1):
        right(c, f"{i}. {note}", right_x - 22, y, 11)
        y -= 28
    y -= 10
    right(c, "ظهر الاستمارة", right_x - 22, bottom + 20, 9, "Arab", Color(0.4, 0.4, 0.4))


def write_pdf(path: Path, page_size: Tuple[float, float]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=page_size)
    c.setTitle("استمارة مخصصات")
    c.setAuthor("نظام الأفراد")
    draw_form(c, page_size[0], page_size[1])
    c.showPage()
    draw_notes_page(c, page_size[0], page_size[1])
    c.save()


def default_outputs() -> list:
    here = Path(__file__).resolve().parent
    outs = [
        here / "makhsasat_new.pdf",
        here / "data" / "makhsasat_new.pdf",
        Path.home() / "Downloads" / "مخصصات_جديدة.pdf",
        Path.home() / "Desktop" / "092026" / "مخصصات_جديدة.pdf",
    ]
    if os.name == "nt":
        outs.extend(
            [
                Path(r"C:\Users\ngc\Downloads\مخصصات_جديدة.pdf"),
                Path(r"C:\Users\ngc\Desktop\092026\مخصصات_جديدة.pdf"),
            ]
        )
    seen = []
    for p in outs:
        if p not in seen:
            seen.append(p)
    return seen


def main() -> int:
    find_font()
    original = find_original()
    if original:
        page_size = read_page_size(original)
        print("الملف الأصلي:")
        print(str(original))
        print("قياس الصفحة نقطة:")
        print(f"{page_size[0]:.2f} x {page_size[1]:.2f}")
    else:
        page_size = A4
        print("لم يُعثر على الملف الأصلي في المسار المعطى.")
        print("سيتم اعتماد قياس")
        print("A4")
        print(f"{page_size[0]:.2f} x {page_size[1]:.2f}")

    written = []
    for out in default_outputs():
        try:
            out.parent.mkdir(parents=True, exist_ok=True)
            write_pdf(out, page_size)
            if out.exists():
                written.append(out)
                print("تم الحفظ:")
                print(str(out))
        except OSError as exc:
            print("تعذر الحفظ في:")
            print(str(out))
            print(str(exc))

    if not written:
        print("لم يُحفظ أي ملف.")
        return 1
    print("الاستمارة جديدة وبفارغة، بنفس القياس والمواصفات.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
