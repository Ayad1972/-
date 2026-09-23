# -*- coding: utf-8 -*-
"""
نسخة طبق الأصل من استمارة مخصصات.pdf
مع تنظيف النقاط السود وإزالة خط اليد وإبقاء النصوص المطبوعة والخطوط.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import List, Optional, Tuple

import numpy as np

try:
    import cv2
except ImportError:
    print("ثبّت: pip install opencv-python-headless numpy pymupdf")
    sys.exit(1)

try:
    import pymupdf
except ImportError:
    print("ثبّت: pip install pymupdf")
    sys.exit(1)


ORIGINAL_CANDIDATES = [
    Path(r"C:\Users\ngc\Desktop\092026\مخصصات.pdf"),
    Path.home() / "Desktop" / "092026" / "مخصصات.pdf",
    Path(__file__).resolve().parent / "data" / "مخصصات.pdf",
    Path(__file__).resolve().parent / "مخصصات.pdf",
]


def find_original(explicit: Optional[str] = None) -> Optional[Path]:
    if explicit:
        p = Path(explicit)
        if p.exists():
            return p
    for p in ORIGINAL_CANDIDATES:
        if p.exists():
            return p
    folder = Path.home() / "Desktop" / "092026"
    if folder.exists():
        for p in folder.glob("*.pdf"):
            if p.name == "مخصصات.pdf":
                return p
    return None


def render_pages(pdf_path: Path, dpi: int = 400) -> Tuple[List[np.ndarray], Tuple[float, float]]:
    doc = pymupdf.open(str(pdf_path))
    if doc.page_count < 1:
        raise ValueError("ملف PDF فارغ")
    zoom = dpi / 72.0
    matrix = pymupdf.Matrix(zoom, zoom)
    pages = []
    rect = doc[0].rect
    page_size = (float(rect.width), float(rect.height))
    for page in doc:
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if pix.n == 4:
            img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
        elif pix.n == 1:
            img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
        else:
            img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
        pages.append(img)
    doc.close()
    return pages, page_size


def drop_colored_ink(bgr: np.ndarray) -> np.ndarray:
    """إبقاء الحبر الأسود المطبوع وإسقاط القلم الملون إن وُجد."""
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
    h, s, v = cv2.split(hsv)
    black = (v < 90) & (s < 80)
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    very_dark = gray < 70
    keep = black | very_dark
    out = np.full_like(gray, 255)
    out[keep] = gray[keep]
    return out


def binarize(gray: np.ndarray) -> np.ndarray:
    blur = cv2.GaussianBlur(gray, (3, 3), 0)
    thr = cv2.adaptiveThreshold(
        blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 12
    )
    # نص أسود على أبيض
    if np.mean(thr) < 127:
        thr = cv2.bitwise_not(thr)
    return thr


def extract_lines(binary_inv: np.ndarray) -> np.ndarray:
    """binary_inv: أبيض=حبر. نخرج الخطوط الأفقية والعمودية الحادة."""
    h, w = binary_inv.shape
    hk = max(25, w // 40)
    vk = max(25, h // 50)
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (hk, 1))
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, vk))
    horiz = cv2.morphologyEx(binary_inv, cv2.MORPH_OPEN, h_kernel)
    vert = cv2.morphologyEx(binary_inv, cv2.MORPH_OPEN, v_kernel)
    lines = cv2.bitwise_or(horiz, vert)
    lines = cv2.dilate(lines, cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2)))
    return lines


def remove_isolated_speckles(binary_inv: np.ndarray, max_speckle: int = 36, near: int = 14) -> np.ndarray:
    """يحذف النقاط السود المعزولة ويبقي نقاط الحروف الملاصقة للنص المطبوع."""
    num, labels, stats, cents = cv2.connectedComponentsWithStats(binary_inv, connectivity=8)
    if num <= 1:
        return binary_inv
    protect = np.zeros_like(binary_inv)
    for i in range(1, num):
        if int(stats[i, cv2.CC_STAT_AREA]) >= 45:
            protect[labels == i] = 255
    if np.count_nonzero(protect):
        k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (near, near))
        protect = cv2.dilate(protect, k)
    keep = np.zeros_like(binary_inv)
    for i in range(1, num):
        area = int(stats[i, cv2.CC_STAT_AREA])
        ys, xs = np.where(labels == i)
        if area >= 45:
            keep[labels == i] = 255
            continue
        if area < 8:
            continue
        # نقطة صغيرة تُحفظ فقط إن لامست حرفاً مطبوعاً
        if np.any(protect[ys, xs] > 0) and area >= 8:
            keep[labels == i] = 255
            continue
        # خطوط قصيرة جداً ضمن الجدول تُحذف إن كانت معزولة
    return keep


def keep_printed_text(ink: np.ndarray, lines: np.ndarray) -> np.ndarray:
    """
    نبقي المكوّنات التي تشبه الحروف المطبوعة.
    نزيل النقاط الصغيرة وخطوط اليد الطويلة غير المنتظمة.
    """
    text_only = cv2.bitwise_and(ink, cv2.bitwise_not(lines))
    num, labels, stats, _ = cv2.connectedComponentsWithStats(text_only, connectivity=8)
    if num <= 1:
        return lines.copy()

    heights = []
    widths = []
    for i in range(1, num):
        h = int(stats[i, cv2.CC_STAT_HEIGHT])
        w = int(stats[i, cv2.CC_STAT_WIDTH])
        area = int(stats[i, cv2.CC_STAT_AREA])
        if 12 <= h <= 90 and 8 <= w <= 220 and area >= 20:
            heights.append(h)
            widths.append(w)
    med_h = float(np.median(heights)) if heights else 28.0
    med_w = float(np.median(widths)) if widths else 20.0

    keep_text = np.zeros_like(ink)
    for i in range(1, num):
        x = int(stats[i, cv2.CC_STAT_LEFT])
        y = int(stats[i, cv2.CC_STAT_TOP])
        w = int(stats[i, cv2.CC_STAT_WIDTH])
        h = int(stats[i, cv2.CC_STAT_HEIGHT])
        area = int(stats[i, cv2.CC_STAT_AREA])
        if area < 12:
            continue  # نقاط سود
        # خطوط طويلة رفيعة أفقية/عمودية تُترك للـ lines
        if w >= 40 and h <= 6:
            continue
        if h >= 40 and w <= 6:
            continue
        aspect = w / float(h) if h else 99
        fill = area / float(max(w * h, 1))
        # خط اليد: شكل ممتد غير منتظم أو أعلى بكثير من ارتفاع الحرف المطبوع
        handwritten = False
        if h > med_h * 2.3 and w > med_w * 2.0:
            handwritten = True
        if h > 110 and aspect > 2.8:
            handwritten = True
        if fill < 0.12 and area > 180:
            handwritten = True
        if aspect > 8 and h > 18:
            handwritten = True
        # حروف مطبوعة عربية تقريباً ضمن نطاق ارتفاع ثابت
        printed = (0.45 * med_h <= h <= 1.85 * med_h) and (w <= 12 * med_w) and fill >= 0.15
        if handwritten and not printed:
            continue
        if h > 140:
            continue
        keep_text[labels == i] = 255

    cleaned = cv2.bitwise_or(lines, keep_text)
    return cleaned


def sharpen_binary(binary_inv: np.ndarray) -> np.ndarray:
    # إغلاق خفيف لملء فراغات الحروف المطبوعة ثم تنحيف النقاط
    k = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    closed = cv2.morphologyEx(binary_inv, cv2.MORPH_CLOSE, k)
    return closed


def clean_page(bgr: np.ndarray) -> np.ndarray:
    gray = drop_colored_ink(bgr)
    bw = binarize(gray)  # 255 أبيض
    ink = cv2.bitwise_not(bw)  # 255 حبر
    ink = remove_isolated_speckles(ink)
    lines = extract_lines(ink)
    cleaned_ink = keep_printed_text(ink, lines)
    cleaned_ink = remove_isolated_speckles(cleaned_ink)
    cleaned_ink = sharpen_binary(cleaned_ink)
    # خلفية بيضاء ناصعة وحبر أسود خالص
    out = np.full_like(cleaned_ink, 255)
    out[cleaned_ink > 0] = 0
    # إطار خفيف إن انكسر
    return cv2.cvtColor(out, cv2.COLOR_GRAY2BGR)


def save_pdf(pages: List[np.ndarray], page_size: Tuple[float, float], dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open()
    w, h = page_size
    for img in pages:
        page = doc.new_page(width=w, height=h)
        ok, encoded = cv2.imencode(".png", img)
        if not ok:
            raise RuntimeError("فشل ترميز الصورة")
        page.insert_image(page.rect, stream=encoded.tobytes())
    doc.save(str(dest), deflate=True)
    doc.close()


def outputs_for(original: Path) -> List[Path]:
    outs = [
        original.with_name("مخصصات_مطبوعة.pdf"),
        Path.home() / "Downloads" / "مخصصات_مطبوعة.pdf",
        Path(__file__).resolve().parent / "makhsasat_printed.pdf",
        Path(__file__).resolve().parent / "data" / "makhsasat_printed.pdf",
    ]
    if os.name == "nt":
        outs.append(Path(r"C:\Users\ngc\Downloads\مخصصات_مطبوعة.pdf"))
        outs.append(Path(r"C:\Users\ngc\Desktop\092026\مخصصات_مطبوعة.pdf"))
    seen: List[Path] = []
    for p in outs:
        if p not in seen:
            seen.append(p)
    return seen


def make_synthetic_scan() -> np.ndarray:
    """اختبار: نموذج مطبوع + نقاط سود + خط يد."""
    img = np.full((900, 700, 3), 255, np.uint8)
    cv2.rectangle(img, (40, 40), (660, 860), (0, 0, 0), 3)
    cv2.rectangle(img, (70, 120), (630, 200), (0, 0, 0), 2)
    cv2.line(img, (70, 280), (630, 280), (0, 0, 0), 2)
    cv2.putText(img, "FORM", (250, 90), cv2.FONT_HERSHEY_SIMPLEX, 1.4, (0, 0, 0), 3)
    cv2.putText(img, "Name", (90, 170), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 0), 2)
    rng = np.random.default_rng(1)
    for _ in range(180):
        x = int(rng.integers(50, 650))
        y = int(rng.integers(50, 850))
        cv2.circle(img, (x, y), int(rng.integers(1, 3)), (0, 0, 0), -1)
    # خط يد منحني
    pts = np.array([[120, 400], [180, 360], [260, 430], [340, 350], [420, 440]], np.int32)
    cv2.polylines(img, [pts], False, (20, 20, 20), 3)
    return img


def self_test() -> int:
    raw = make_synthetic_scan()
    cleaned = clean_page(raw)
    gray = cv2.cvtColor(cleaned, cv2.COLOR_BGR2GRAY)
    ink = ((gray < 80).astype(np.uint8)) * 255
    num, labels, stats, _ = cv2.connectedComponentsWithStats(ink, connectivity=8)
    leftover = 0
    for i in range(1, num):
        area = int(stats[i, cv2.CC_STAT_AREA])
        w = int(stats[i, cv2.CC_STAT_WIDTH])
        h = int(stats[i, cv2.CC_STAT_HEIGHT])
        if area <= 36 and w <= 9 and h <= 9:
            leftover += 1
    print("SELF-TEST leftover isolated dots", leftover)
    if leftover > 8:
        raise AssertionError("النقاط السود لم تُنظف كفاية")
    dest = Path("/tmp/makhsasat_self_test.pdf")
    save_pdf([cleaned], (595.27, 841.89), dest)
    print("SELF-TEST OK", dest)
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    args = argv if argv is not None else sys.argv[1:]
    if args and args[0] == "--self-test":
        return self_test()
    explicit = args[0] if args else None
    original = find_original(explicit)
    if original is None:
        print("لم يُعثر على الاستمارة الأصلية.")
        print("ضع الملف في:")
        print(r"C:\Users\ngc\Desktop\092026\مخصصات.pdf")
        print("ثم شغّل run_clean_makhsasat.bat")
        return 1

    print("الأصل:")
    print(str(original))
    pages, page_size = render_pages(original, dpi=400)
    print("عدد الصفحات:", len(pages))
    print("قياس الصفحة:", f"{page_size[0]:.2f} x {page_size[1]:.2f}")

    cleaned_pages = []
    for i, page in enumerate(pages, 1):
        print("تنظيف الصفحة", i)
        cleaned_pages.append(clean_page(page))

    written = []
    for dest in outputs_for(original):
        try:
            save_pdf(cleaned_pages, page_size, dest)
            written.append(dest)
            print("تم الحفظ:")
            print(str(dest))
        except OSError as exc:
            print("تعذر الحفظ في:")
            print(str(dest))
            print(str(exc))

    if not written:
        return 1
    print("النتيجة: نسخة مطبوعة واضحة بلا نقاط سود وبلا خط يد، بنفس قياس الأصل.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
