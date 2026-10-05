#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MAJID EISAPOUR — Design & Spatial Practice
Project : Backyard Landscape — Qaemshahr, Mazandaran
Document: Landscape Design & Technical BOQ  (دفترچه طراحی منظر و برآورد فنی)

Builder script — این اسکریپت تنها منبع ساخت فایل اکسل است.
هر تغییری در محتوا یا گرافیک، از همین‌جا انجام می‌شود و فایل به‌صورت
قطعی و تکرارپذیر (deterministic) بازتولید می‌گردد.

Run:  python3 tools/build_workbook.py
"""

from __future__ import annotations

import datetime as _dt
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.data_source import AxDataSource
from openpyxl.chart.series import StrRef
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties,
)
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.fills import GradientFill, Stop
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(REPO, "MAJID_EISAPOUR_QAEMSHAHR_LANDSCAPE_BOQ_PREMIUM.xlsx")

REV = "02"
DOC_DATE_EN = "1405/07 · Oct 2026"

# ---------------------------------------------------------------------------
# 01 | DESIGN SYSTEM  —  پالت، فونت و ریتم بصری
# ---------------------------------------------------------------------------
INK      = "1E2A26"   # deep forest charcoal  (عنوان‌ها، بندهای تیره)
INK_2    = "33413B"   # ink, one step lighter (سرستون، بندهای فرعی)
SAGE     = "7C8B72"   # sage green            (لهجه سبز)
SAGE_DK  = "59684F"   # sage, dark            (متن تأکیدی)
SAGE_LT  = "E7EBDF"   # sage, wash            (بند بیانیه)
TERRA    = "B4713F"   # terracotta            (خط تأکید / اعداد)
TERRA_LT = "F6E7DA"
BRASS    = "C0A163"   # brass                 (متن روی زمینه تیره)
CREAM    = "F8F5EC"   # base paper
CREAM_2  = "F1EBDD"   # alt row / label
WHITE    = "FFFFFF"
LINE     = "D8D1C0"
LINE_LT  = "EAE4D6"
AMBER    = "FBEFD6"   # editable-input wash
AMBER_TX = "8A5A16"
GOOD     = "3F6B4A"   # OK / control
BAD      = "A3352C"

FD  = "Aptos Display"   # latin display
FB  = "Aptos"           # latin body
FP  = "Tahoma"          # persian  (pre-installed, correct Farsi rendering)

MONEY = '#,##0'
QTY   = '#,##0.##'
PCT   = '0.0%'

import unicodedata as _ud


def has_fa(text) -> bool:
    """True when the first strong directional character is RTL (Persian/Arabic).

    Excel's General alignment follows the paragraph direction, and the paragraph
    direction follows the first strong character — digits and symbols are skipped.
    """
    for ch in str(text or ""):
        bidi = _ud.bidirectional(ch)
        if bidi in ("R", "AL"):
            return True
        if bidi == "L":
            return False
    return False


def align(value=None, h="general", v="center", wrap=True, ro=None):
    if ro is None:
        ro = 2 if has_fa(value) else 1
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, readingOrder=ro)


def fnt(size=9.5, bold=False, italic=False, color=INK, name=FP):
    return Font(name=name, size=size, bold=bold, italic=italic, color=color)


def sol(color):
    return PatternFill("solid", fgColor=color)


def bd(color=LINE, style="thin"):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)


NO_BORDER = Border()


# ---------------------------------------------------------------------------
# 02 | CELL ROLES  —  نقش‌های بصری مشترک بین همه برگه‌ها
# ---------------------------------------------------------------------------
ROLES = {
    # ---- bands -------------------------------------------------------------
    "band_en":    dict(font=fnt(24, True, color=WHITE, name=FD), fill=sol(INK),
                       align=align(h="center", wrap=False)),
    "band_fa":    dict(font=fnt(12.5, True, color=CREAM, name=FP), fill=sol(INK),
                       align=align(h="center", wrap=False)),
    "band_meta":  dict(font=fnt(9, False, color=BRASS, name=FB), fill=sol(INK),
                       align=align(h="center", wrap=False)),
    "kicker":     dict(font=fnt(9, False, color=SAGE_DK, name=FB, italic=True),
                       fill=sol(CREAM), align=align(h="center")),
    "rule":       dict(fill=sol(TERRA)),
    "rule_gold":  dict(fill=sol(BRASS)),
    "spacer":     dict(fill=sol(CREAM)),
    "colhead":    dict(font=fnt(9.5, True, color=WHITE, name=FP), fill=sol(INK_2),
                       align=align(h="center"), border=bd(color=LINE)),
    # ---- rows --------------------------------------------------------------
    "row_text":   dict(font=fnt(9), fill=sol(WHITE), align=align(h="general"),
                       border=bd()),
    "row_text_a": dict(font=fnt(9), fill=sol(CREAM_2), align=align(h="general"),
                       border=bd()),
    "row_code":   dict(font=fnt(9, True, color=SAGE_DK), fill=sol(WHITE),
                       align=align(h="center"), border=bd()),
    "row_code_a": dict(font=fnt(9, True, color=SAGE_DK), fill=sol(CREAM_2),
                       align=align(h="center"), border=bd()),
    "row_num":    dict(font=fnt(9, color=INK_2), fill=sol(WHITE),
                       align=align(h="center"), border=bd()),
    "row_num_a":  dict(font=fnt(9, color=INK_2), fill=sol(CREAM_2),
                       align=align(h="center"), border=bd()),
    "row_money":  dict(font=fnt(9, color=INK, name=FB), fill=sol(WHITE),
                       align=align(h="center"), border=bd()),
    "row_money_a": dict(font=fnt(9, color=INK, name=FB), fill=sol(CREAM_2),
                        align=align(h="center"), border=bd()),
    "row_note":   dict(font=fnt(8.5, color=SAGE_DK), fill=sol(WHITE),
                       align=align(h="general"), border=bd()),
    "row_note_a": dict(font=fnt(8.5, color=SAGE_DK), fill=sol(CREAM_2),
                       align=align(h="general"), border=bd()),
    # ---- boq grouping ------------------------------------------------------
    "secbar":     dict(font=fnt(9.5, True, color=INK_2), fill=sol(SAGE_LT),
                       align=align(h="general"),
                       border=Border(bottom=Side("medium", TERRA))),
    "subtotal":   dict(font=fnt(9.5, True, color=INK), fill=sol(CREAM_2),
                       align=align(h="general"),
                       border=Border(top=Side("thin", SAGE), bottom=Side("thin", LINE))),
    "subtotal_v": dict(font=fnt(9.5, True, color=INK, name=FB), fill=sol(CREAM_2),
                       align=align(h="center"),
                       border=Border(top=Side("thin", SAGE), bottom=Side("thin", LINE))),
    "grand":      dict(font=fnt(10.5, True, color=WHITE, name=FP), fill=sol(INK),
                       align=align(h="general"), border=bd(color=INK)),
    "grand_v":    dict(font=fnt(10.5, True, color=WHITE, name=FB), fill=sol(INK),
                       align=align(h="center"), border=bd(color=INK)),
    "grand_alt":  dict(font=fnt(10.5, True, color=INK), fill=sol(CREAM_2),
                       align=align(h="general"), border=bd(color=SAGE)),
    "grand_alt_v": dict(font=fnt(10.5, True, color=INK, name=FB), fill=sol(CREAM_2),
                        align=align(h="center"), border=bd(color=SAGE)),
    "input":      dict(font=fnt(10, True, color=AMBER_TX, name=FB), fill=sol(AMBER),
                       align=align(h="center"),
                       border=Border(*(Side("dashed", BRASS),) * 4)),
    # ---- info / narrative --------------------------------------------------
    "label":      dict(font=fnt(9, True, color=INK), fill=sol(CREAM_2),
                       align=align(h="general"), border=bd(color=LINE_LT)),
    "value":      dict(font=fnt(9, color=INK_2), fill=sol(WHITE),
                       align=align(h="general"), border=bd(color=LINE_LT)),
    "mini":       dict(font=fnt(10, True, color=INK), fill=sol(SAGE_LT),
                       align=align(h="general"),
                       border=Border(bottom=Side("thin", SAGE))),
    "quote":      dict(font=fnt(10.5, italic=True, color=INK, name=FP),
                       fill=sol(SAGE_LT), align=align(h="center")),
    "body":       dict(font=fnt(9.5, color=INK_2), fill=sol(CREAM),
                       align=align(h="general")),
    "contents":   dict(font=fnt(9.5, color=INK_2), fill=sol(WHITE),
                       align=align(h="general"), border=bd(color=LINE_LT)),
    "contents_a": dict(font=fnt(9.5, color=INK_2), fill=sol(CREAM_2),
                       align=align(h="general"), border=bd(color=LINE_LT)),
    "code_cell":  dict(font=fnt(9.5, True, color=WHITE, name=FB), fill=sol(INK_2),
                       align=align(h="center"), border=bd(color=LINE)),
    "note_line":  dict(font=fnt(8.5, italic=True, color=SAGE_DK), fill=sol(CREAM),
                       align=align(h="general")),
    "kpi_val":    dict(font=fnt(12, True, color=TERRA, name=FB), fill=sol(WHITE),
                       align=align(h="center"), border=bd(color=LINE)),
}


def put(ws, ref, value=None, role=None, *, merge=None, fmt=None, h=None, v=None,
        fill=None, font=None, align_=None, border=None, link=None):
    """Write one styled cell (and optionally merge a range)."""
    cell = ws[ref]
    if value is not None:
        cell.value = value
    spec = ROLES.get(role, {})
    cell.font = font or spec.get("font", fnt(9))
    cell.fill = fill or spec.get("fill", sol(CREAM))
    a = spec.get("align", align(value))
    if h or v or align_:
        a = align_(value) if align_ else align(value, h=h or a.horizontal,
                                                v=v or a.vertical)
    cell.alignment = a
    border_ob = NO_BORDER
    if border is not None:
        border_ob = border
    elif spec.get("border") is not None:
        border_ob = spec["border"]
    cell.border = border_ob
    if fmt:
        cell.number_format = fmt
    if link:
        cell.hyperlink = link
    if merge:
        first, last_ = merge.split(":") if ":" in merge else (merge, merge)
        ws.merge_cells(f"{first}:{last_}")
        # borders must be painted on every cell of a merged block
        if any(getattr(sd, "style", None) for sd in (border_ob.left, border_ob.right,
                                                     border_ob.top, border_ob.bottom)):
            for row in ws[f"{first}:{last_}"]:
                for c in row:
                    c.border = border_ob
    return cell


def paint(ws, ref_range, role, border_role=None):
    """Fill/font a whole range (used for merged bands and rules)."""
    spec = ROLES.get(role, {})
    br = ROLES.get(border_role, {}).get("border") if border_role else None
    for row in ws[ref_range]:
        for c in row:
            c.fill = spec.get("fill", sol(CREAM))
            if spec.get("font"):
                c.font = spec["font"]
            if spec.get("align"):
                c.alignment = spec["align"]
            if br:
                c.border = br


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def heights(ws, mapping):
    for r, h in mapping.items():
        ws.row_dimensions[r].height = h


def band(ws, ncols, doc_code, title_en, title_fa, kicker, meta="", meta2=""):
    """Standard page head: title band + kicker + hairline rule + column row."""
    last = get_column_letter(ncols)
    heights(ws, {1: 38, 2: 22, 3: 19, 4: 3.5})
    if meta or meta2:
        t_last = get_column_letter(ncols - 3 if ncols >= 6 else 2)
        m_first = get_column_letter(ncols - 2 if ncols >= 6 else 3)
        put(ws, "A1", title_en, "band_en", merge=f"A1:{t_last}1",
            align_=lambda v: align(v, h="general", wrap=False))
        put(ws, "A2", title_fa, "band_fa", merge=f"A2:{t_last}2")
        put(ws, f"{m_first}1", meta or doc_code, "band_meta", merge=f"{m_first}1:{last}1")
        put(ws, f"{m_first}2", meta2, "band_meta", merge=f"{m_first}2:{last}2")
    else:
        put(ws, "A1", title_en, "band_en", merge=f"A1:{last}1")
        put(ws, "A2", title_fa, "band_fa", merge=f"A2:{last}2")
    put(ws, "A3", kicker, "kicker", merge=f"A3:{last}3")
    paint(ws, f"A4:{last}4", "rule")


def colheads(ws, row, headers, height=30):
    heights(ws, {row: height})
    for i, htext in enumerate(headers, start=1):
        put(ws, f"{get_column_letter(i)}{row}", htext, "colhead")


def foot_strip(ws, row, ncols, text):
    last = get_column_letter(ncols)
    heights(ws, {row: 16})
    put(ws, f"A{row}", text, "band_meta", merge=f"A{row}:{last}{row}",
        fill=sol(INK), font=fnt(8, color=CREAM, name=FB),
        align_=lambda v: align(v, h="center", wrap=False))


def finish(ws, ncols, *, freeze=None, titles=None, header=True, fit_h=0,
           tab=None, zoom=100, footer_left=None, footer_right=None):
    last = get_column_letter(ncols)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.zoomScale = zoom
    if freeze:
        ws.freeze_panes = freeze
    if tab:
        ws.sheet_properties.tabColor = tab
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9            # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_h
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.35
    ws.page_margins.top = ws.page_margins.bottom = 0.45
    ws.page_margins.header = ws.page_margins.footer = 0.2
    ws.print_area = f"A1:{last}{ws.max_row}"
    if titles:
        ws.print_title_rows = titles
    if header:
        ws.oddHeader.left.text = "MAJID EISAPOUR   ·   Design && Spatial Practice"
        ws.oddHeader.left.size = 8
        ws.oddHeader.left.color = "59684F"
        ws.oddHeader.right.text = (
            f"BACKYARD LANDSCAPE · QAEMSHAHR   |   {footer_left or ''}   |   REV {REV}"
        )
        ws.oddHeader.right.size = 8
        ws.oddHeader.right.color = "59684F"
        ws.oddFooter.left.text = footer_right or "Landscape Design && Technical BOQ  —  دفترچه طراحی منظر و برآورد فنی"
        ws.oddFooter.left.size = 8
        ws.oddFooter.left.color = "7C8B72"
        ws.oddFooter.right.text = "&A   ·   &P / &N"
        ws.oddFooter.right.size = 8
        ws.oddFooter.right.color = "7C8B72"


# ---------------------------------------------------------------------------
# 03 | CONTENT  —  محتوای پروژه (منبع واحد حقیقت)
# ---------------------------------------------------------------------------
PROJECT = dict(
    title_fa="حیاط پشتی ویلای مسکونی",
    title_en="Residential Backyard Villa",
    place_fa="قائمشهر، مازندران",
    place_en="Qaemshahr, Mazandaran",
    area="≈ 495 m²",
    client="—",
    owner_note="تکمیل توسط کارفرما",
)

NARRATIVE = [
    ("01", "ایده مرکزی / Core Idea", "OUTDOOR ROOM — اتاقِ بیرونی",
     "The garden is conceived as an extension of domestic life, not as a collection of isolated objects.\n"
     "باغ به‌عنوان امتداد زندگی روزمره طراحی شده است، نه مجموعه‌ای از عناصر جدا از هم."),
    ("02", "توالی ورود / Arrival Sequence", "FROM PARKING TO GARDEN — از مسیر به باغ",
     "The entrance begins directly from the parking corridor; the path becomes a spatial threshold through "
     "texture, planting and controlled views.\n"
     "ورودی مستقیماً از مسیر پارکینگ آغاز می‌شود و با تغییر بافت، کاشت و کنترل دید به یک آستانه فضایی تبدیل می‌شود."),
    ("03", "قلب سبز / Green Heart", "OPEN LAWN — میدان سبز",
     "A quiet central clearing balances the hardscape and gives the garden visual breathing room.\n"
     "فضای سبز مرکزی، تعادل میان کف‌سازی و کاشت را برقرار کرده و به باغ فرصت تنفس بصری می‌دهد."),
    ("04", "فضای زندگی / Living Space", "PAVILION + TERRACE — پاویون و تراس",
     "A compact architectural pavilion and open terrace form one continuous social composition.\n"
     "پاویون و تراس باز به‌عنوان یک ترکیب پیوسته برای نشستن، گفت‌وگو و مکث شکل می‌گیرند."),
    ("05", "آب معماری / Architectural Water", "LINEAR WATER — آبنمای خطی",
     "A narrow linear water element accompanies circulation as a reflective architectural line rather than a "
     "standalone fountain.\n"
     "آب به‌صورت یک خط باریک و بازتابنده در کنار مسیر حرکت قرار می‌گیرد، نه به‌عنوان آبنمایی مستقل و جداافتاده."),
]

# کد | موضوع | مبنای طراحی | وضعیت | واحد | مقدار | یادداشت | توضیح
BASIS = [
    ("B-01", "مساحت سایت / Site Area", "20.2 × 24.5 = 494.9", "مبنای محاسبه / Basis",
     "m²", 495, "Conceptual", "فقط حیاط پشتی / Backyard only"),
    ("B-02", "پاویون / Pavilion", "5.40 × 4.80", "طراحی / Design",
     "m²", 25.9, "Indicative", "طرح سازه نهایی توسط مهندس سازه"),
    ("B-03", "تراس باز / Open Terrace", "6.60 × 3.20 ≈ 21.1", "طراحی / Design",
     "m²", 21.1, "Indicative", "یکپارچه با پاویون / Integrated with pavilion"),
    ("B-04", "آبنمای خطی / Linear Water", "11.00 × 0.50", "طراحی / Design",
     "m²", 5.5, "Indicative", "عنصر بازتابنده و معماری / Reflective element"),
    ("B-05", "باغ سبزی / Kitchen Garden", "زون 120 m² = 90 بستر + 30 دسترسی", "طراحی / Design",
     "m²", 120, "Seasonal", "بستر خالص کاشت 90 m² (تطبیق با BOQ 09-01 و P-05)"),
    ("B-06", "لبه خدماتی / Service Edge", "Storage + washing + lobby", "موجود / Existing",
     "—", "—", "Untouched", "بدون طراحی مجدد / No redesign included"),
    ("B-07", "ورودی / Entry", "Direct from parking corridor", "موجود / Existing",
     "—", "—", "Integrated", "بدون درب/گیت فرض شده / No gate assumed"),
    ("B-08", "کفپوش سخت / Hardscape Paving", "مسیر 45 + تراس 22 = 67", "طراحی / Design",
     "m²", 67, "Indicative", "تطبیق با BOQ 03-01 و 05-01 و M-01"),
    ("B-09", "چمن / Lawn", "قلب سبز مرکزی / Central green heart", "طراحی / Design",
     "m²", 140, "Indicative", "تطبیق با BOQ 08-01 و P-06"),
    ("B-10", "تفکیک سطوح / Area Breakdown", "67 سخت + 140 چمن + 120 باغ سبزی + ≈168 کاشت و لبه",
     "کنترلی / Control", "m²", 495, "Indicative", "جمع کنترلی با B-01 / Reconciles with B-01"),
]

# کد | بخش و آیتم | شرح فنی | واحد | مقدار | قیمت واحد | یادداشت و ارجاعات
BOQ_SECTIONS = [
    ("01", "تجهیز کارگاه و حفاظت / SITE SETUP & PROTECTION", [
        ("01-01", "تجهیز کارگاه / Site Setup",
         "محافظت از محدوده موجود، ابزار و تجهیز اولیه", "مقطوع (LS)", 1, 35_000_000,
         "شروع پروژه؛ ورود به لبه خدماتی ممنوع (→ B-06)"),
    ]),
    ("02", "عملیات خاکی و آماده‌سازی بستر / EARTHWORK & SOIL PREPARATION", [
        ("02-01", "خاک‌برداری / Excavation",
         "برداشت موضعی برای زیرسازی مسیرها و سازه", "m³", 35, 450_000,
         "برداشت موضعی و کنترل‌شده (→ B-02، B-03)"),
        ("02-02", "خاک زراعی / Topsoil",
         "تأمین، اصلاح و پخش خاک زراعی", "m³", 28, 700_000,
         "بستر کاشت و باغ سبزی (→ M-06)"),
        ("02-03", "کمپوست / Compost",
         "ماده آلی و اصلاح خاک", "m³", 8, 950_000,
         "اختلاط کنترل‌شده (→ M-06، TN-07)"),
    ]),
    ("03", "کف‌پوش و محوطه سخت / PAVING & HARDSCAPE", [
        ("03-01", "کف‌سازی سنگ / Stone Paving",
         "سنگ طبیعی، برش، ملات/بستر و بندکشی", "m²", 45, 3_200_000,
         "مسیرها؛ سنگ منتخب (→ M-01)"),
        ("03-02", "زیرسازی / Subbase",
         "مصالح دانه‌بندی‌شده + تراکم مرحله‌ای", "m³", 28, 900_000,
         "تراکم مرحله‌ای الزامی است (→ TN-01)"),
    ]),
    ("04", "پاویون / PAVILION", [
        ("04-01", "فونداسیون پاویون / Foundation",
         "بتن، آرماتور و قالب‌بندی اولیه", "m²", 30, 8_500_000,
         "ضخامت مبنا ≈20 سانتی‌متر ≈ 6 m³ بتن (→ M-04، TN-03)"),
        ("04-02", "اسکلت فلزی / Steel Frame",
         "پروفیل، برش، جوش، ضدزنگ و نصب", "kg", 1200, 210_000,
         "وزن تقریبی؛ تأیید مهندس سازه (→ M-03)"),
        ("04-03", "سقف / Roof",
         "پوسته سقف و لایه‌های آب‌بندی", "m²", 32, 4_200_000,
         "هماهنگ با دیتیل معماری (→ B-02)"),
        ("04-04", "ترمووود / Thermowood",
         "سقف کاذب و پوسته چوبی منتخب", "m²", 30, 3_500_000,
         "وابسته به گونه و گرید (→ M-02، TN-04)"),
    ]),
    ("05", "تراس باز / OPEN TERRACE", [
        ("05-01", "تراس / Terrace",
         "کف‌سازی و لبه‌سازی فضای نشستن", "m²", 22, 3_000_000,
         "یکپارچه با پاویون (→ B-03، M-01)"),
    ]),
    ("06", "آبنمای خطی / LINEAR WATER FEATURE", [
        ("06-01", "آبنمای خطی / Linear Water",
         "کاسه، آب‌بندی، لبه سنگی و گردش بسته آب", "lm", 11, 15_000_000,
         "آبنمای کم‌عمق و معماری (→ B-04، M-05، TN-06)"),
    ]),
    ("07", "آبیاری / IRRIGATION", [
        ("07-01", "آبیاری / Irrigation Network",
         "شبکه اصلی، تفکیک زون و کنترل", "m²", 495, 190_000,
         "شامل شبکه اصلی ≈120 lm و 6 عدد شیر برقی (→ I-01، I-03، TN-08)"),
        ("07-02", "لوله قطره‌ای / Drip Line",
         "لوله و اتصالات قطره‌ای", "lm", 350, 100_000,
         "ویژه باغ سبزی و کاشت (→ I-02)"),
    ]),
    ("08", "کاشت و چمن / PLANTING & LAWN", [
        ("08-01", "چمن / Lawn",
         "زیرسازی و اجرای چمن", "m²", 140, 270_000,
         "چمن مقاوم و کم‌دردسر (→ P-06)"),
        ("08-02", "درخت شاخص / Specimen Tree",
         "درخت شاخص با کاشت و تثبیت", "عدد", 2, 12_000_000,
         "انتخاب نهایی با توجه به اقلیم (→ P-01)"),
        ("08-03", "درخت میوه / Fruit Tree",
         "نهال میوه مناسب منطقه", "عدد", 4, 5_500_000,
         "در محدوده باغ تولیدی (→ P-02)"),
        ("08-04", "درختچه و پرچین / Shrub & Hedge",
         "کاشت درختچه چندلایه و پرچین", "عدد", 35, 950_000,
         "شامل پرچین 25 متر طول (→ P-03، P-04)"),
    ]),
    ("09", "باغ سبزی / KITCHEN GARDEN", [
        ("09-01", "باغچه تولیدی / Raised Beds",
         "ساخت بسترهای کاشت سبزی", "m²", 90, 1_500_000,
         "بستر خالص در زون 120 m² (→ B-05، P-05)"),
    ]),
    ("10", "روشنایی محوطه / LIGHTING", [
        ("10-01", "چراغ مسیر / Path Light",
         "چراغ محوطه با IP مناسب", "عدد", 10, 4_800_000,
         "نور کم‌ارتفاع و کنترل‌شده (→ L-01، TN-05)"),
        ("10-02", "چراغ توکار / Inground Light",
         "نورپردازی تأکیدی معماری", "عدد", 8, 5_200_000,
         "بدون خیرگی؛ درین مناسب (→ L-02)"),
        ("10-03", "چراغ خطی آب / Water Linear Light",
         "نور خطی ضدآب", "lm", 10, 1_200_000,
         "هماهنگ با دیتیل آب (→ L-03)"),
    ]),
    ("11", "کنترل کیفیت و تحویل / QA & HANDOVER", [
        ("11-01", "کنترل کیفیت و تحویل / QA & Handover",
         "بازدید، رفع نواقص و تحویل اولیه", "مقطوع (LS)", 1, 65_000_000,
         "دو مرحله کنترل کیفیت (→ TN-09)"),
    ]),
]

MATERIALS = [
    ("M-01", "سنگ طبیعی / Natural Stone", "مسیرها و تراس", "Calm • tactile • durable",
     "m²", "≈ 67", "جمع 45 مسیر + 22 تراس؛ نمونه و گرید نهایی پیش از سفارش تأیید شود (→ BOQ 03-01، 05-01)"),
    ("M-02", "ترمووود / Thermowood", "سقف و جزئیات پاویون", "Warm • precise • natural",
     "m²", "≈ 30", "زیرسازی و فاصله تهویه پشت چوب الزامی است (→ TN-04)"),
    ("M-03", "فلز رنگ‌شده / Coated Steel", "اسکلت پاویون", "Light • minimal • architectural",
     "kg", "≈ 1,200", "ضدخوردگی مناسب رطوبت شمال (→ BOQ 04-02)"),
    ("M-04", "بتن / Concrete", "فونداسیون و بسترها", "Mass • quiet • robust",
     "m³", "≈ 6", "معادل پی 30 m² با ضخامت ≈20 cm؛ عیار توسط مهندس (→ BOQ 04-01)"),
    ("M-05", "آب / Water", "آبنمای خطی", "Reflective • sensory • calm",
     "lm", "≈ 11", "گردش بسته و دسترسی سرویس (→ BOQ 06-01)"),
    ("M-06", "خاک اصلاح‌شده و کمپوست / Amended Soil & Compost", "بستر کاشت و باغ سبزی",
     "Living • fertile • local", "m³", "≈ 36", "جمع 28 خاک زراعی + 8 کمپوست (→ BOQ 02-02، 02-03)"),
]

PLANTING = [
    ("P-01", "درخت شاخص / Specimen", "Focal tree / نقطه کانونی", "مرکز باغ", "عدد", "2",
     "فرم معماری، سایه ملایم، سازگار با اقلیم (→ BOQ 08-02)"),
    ("P-02", "درخت میوه / Fruit", "Productive layer / لایه تولیدی", "شمال", "عدد", "4",
     "گونه‌های سازگار با اقلیم و نگهداری واقعی (→ BOQ 08-03)"),
    ("P-03", "درختچه / Shrub", "Structure + texture / ساختار و بافت", "حاشیه‌ها", "عدد", "35",
     "ترکیب همیشه‌سبز و فصلی (→ BOQ 08-04)"),
    ("P-04", "پرچین / Hedge", "Filter + frame / فیلتر و قاب", "لبه‌ها", "lm", "25",
     "ارتفاع نهایی کنترل‌شده؛ در ردیف 08-04 لحاظ شده است"),
    ("P-05", "سبزی فصلی / Seasonal Vegetable", "Edible landscape / منظر خوراکی", "باغ سبزی شمال",
     "m²", "90", "بستر خالص از زون 120 m² باغ سبزی (→ B-05، BOQ 09-01)"),
    ("P-06", "چمن / Lawn", "Visual breathing room / تنفس بصری", "مرکز", "m²", "140",
     "سطح یکدست و مقاوم (→ BOQ 08-01)"),
]

LIGHTING_IRRIGATION = [
    ("L-01", "Lighting", "چراغ مسیر / Path Light", "عدد", "10",
     "Low glow / نور کم‌ارتفاع", "IP مناسب فضای باز؛ خیرگی کنترل شود (→ BOQ 10-01)"),
    ("L-02", "Lighting", "چراغ توکار / Inground", "عدد", "8",
     "Accentuate architecture / تأکید معماری", "دسترسی سرویس و درین مناسب (→ BOQ 10-02)"),
    ("L-03", "Lighting", "نور خطی آب / Water Light", "lm", "10",
     "Continuous reflection / بازتاب پیوسته", "کاملاً مناسب محیط مرطوب (→ BOQ 10-03)"),
    ("I-01", "Irrigation", "شبکه اصلی / Main Line", "lm", "120",
     "Efficient distribution / توزیع کارا", "تفکیک زون‌های آبیاری (→ BOQ 07-01)"),
    ("I-02", "Irrigation", "خط قطره‌ای / Drip Line", "lm", "350",
     "Water efficiency / بازده آبی", "ویژه باغ سبزی و کاشت (→ BOQ 07-02)"),
    ("I-03", "Irrigation", "شیر برقی / Solenoid Valve", "عدد", "6",
     "Zone control / کنترل زون", "دسترسی برای تعمیرات (→ BOQ 07-01)"),
]

NOTES = [
    ("TN-01", "زیرسازی / Subbase",
     "تمام سطوح سخت پیش از اجرای لایه نهایی باید متراکم، تراز و دارای شیب کنترل‌شده باشند (→ BOQ 03-02)."),
    ("TN-02", "زهکشی / Drainage",
     "در اقلیم مرطوب، مسیر آب سطحی باید پیش از کف‌سازی حل شود؛ آب نباید به سازه پاویون هدایت شود."),
    ("TN-03", "پاویون / Pavilion",
     "ابعاد این دفتر مفهومی است؛ نقشه سازه، فونداسیون، اتصالات و جزئیات نهایی توسط مهندس ذی‌صلاح تأیید شود (→ BOQ 04)."),
    ("TN-04", "چوب / Timber",
     "ترمووود باید از تأمین‌کننده معتبر و با گرید مناسب فضای بیرونی تهیه شود؛ زیرسازی و تهویه پشت چوب الزامی است (→ M-02)."),
    ("TN-05", "نور / Lighting",
     "چراغ‌های فضای باز باید برای محیط مرطوب و بارندگی انتخاب شوند؛ مسیر کابل و دسترسی تعمیرات از ابتدا پیش‌بینی شود (→ BOQ 10)."),
    ("TN-06", "آب / Water",
     "آبنمای خطی باید دارای گردش بسته، دسترسی سرویس، آب‌بندی مناسب و امکان تخلیه باشد (→ BOQ 06-01)."),
    ("TN-07", "کاشت / Planting",
     "گونه‌ها بر اساس رطوبت، زهکشی، نور، نگهداری و دسترسی محلی نهایی شوند؛ این فهرست جایگزین انتخاب گیاه توسط مجری نیست."),
    ("TN-08", "لبه خدماتی / Existing Service Edge",
     "انبار، فضای شست‌وشوی نیمه‌باز و لابی موجود در دامنه طراحی مجدد نیستند و باید در زمان اجرا حفظ شوند (→ B-06)."),
    ("TN-09", "مبنای برآورد / Estimating Basis",
     "تمام مقادیر از مبانی برگه 03 استخراج شده‌اند؛ پیش از اجرا برداشت نهایی (Take-off) و استعلام قیمت انجام شود. "
     "ردیف‌های مربوط به پرچین و شیر برقی در آیتم‌های 08-04 و 07-01 لحاظ شده‌اند. قیمت‌ها بدون احتساب مالیات بر ارزش افزوده است (→ BOQ)."),
]

REVISIONS = [
    ("01", "1405 / 2026", "نسخه پایه BOQ و ساختار حرفه‌ای دو زبانه",
     "MAJID EISAPOUR", "—", "—"),
    ("02", "1405/07 · Oct 2026",
     "بازطراحی گرافیکی و صفحه‌آرایی، اصلاح انسجام مقادیر بین برگه‌ها، "
     "افزودن ردیف‌های جمع بخش، شماره‌گذاری بخش‌ها (12 ← 11)، بازنویسی فرمول‌ها، "
     "اصلاح محدوده چاپ و تنظیمات پرینت",
     "MAJID EISAPOUR", "—", "—"),
    ("03", "—", "—", "—", "—", "—"),
]

CONTENTS = [
    ("01", "شناسنامه پروژه / Project Information", "01 | Project Info"),
    ("02", "روایت و ایده طراحی / Design Narrative", "02 | Design Narrative"),
    ("03", "مبنای طراحی و دامنه کار / Design Basis & Scope", "03 | Design Basis"),
    ("04", "جدول مقادیر و برآورد / Bill of Quantities", "04 | BOQ | برآورد"),
    ("05", "جدول متریال / Material Schedule", "05 | Material Schedule"),
    ("06", "پالت کاشت / Planting Schedule", "06 | Planting Schedule"),
    ("07", "نورپردازی و آبیاری / Lighting & Irrigation", "07 | Lighting & Irrigation"),
    ("08", "یادداشت‌های اجرایی / Technical Notes & QA", "08 | Technical Notes"),
    ("09", "خلاصه هزینه / Cost Summary & Control", "09 | Cost Summary"),
    ("10", "تاریخچه نسخه‌ها / Revision Log", "10 | Revision Log"),
]


# ---------------------------------------------------------------------------
# 04 | SHEET BUILDERS
# ---------------------------------------------------------------------------
def build_cover(wb):
    ws = wb.create_sheet("00 | جلد و هویت")
    n = 8
    last = get_column_letter(n)
    widths(ws, {c: 13.5 for c in "ABCDEFGH"})
    heights(ws, {
        1: 30, 2: 26, 3: 3.5, 4: 8, 5: 36, 6: 26, 7: 18, 8: 3.5, 9: 10,
        10: 22, 11: 24, 12: 24, 13: 24, 14: 24, 15: 10, 16: 22,
    })

    # --- monogram + studio band
    put(ws, "A1", "ME", merge="A1:B2", fill=GradientFill(
        stop=[Stop(INK_2, 0), Stop(INK, 1)], degree=0),
        font=fnt(30, True, color=BRASS, name=FD),
        align_=lambda v: align(v, h="center", v="center", wrap=False),
        border=Border(bottom=Side("thin", BRASS), right=Side("thin", BRASS)))
    put(ws, "C1", "MAJID EISAPOUR", "band_en", merge=f"C1:{last}1")
    put(ws, "C2", "Design & Spatial Practice   |   استودیو طراحی و فضا",
        "band_meta", merge=f"C2:{last}2",
        font=fnt(10.5, color=BRASS, name=FP))
    paint(ws, f"A3:{last}3", "rule")

    # --- document title
    put(ws, "A5", "LANDSCAPE DESIGN & TECHNICAL BOQ", "band_en", merge=f"A5:{last}5",
        fill=sol(CREAM), font=fnt(21, True, color=INK, name=FD))
    put(ws, "A6", "دفترچه طراحی منظر و برآورد فنی", "band_en", merge=f"A6:{last}6",
        fill=sol(CREAM), font=fnt(14.5, True, color=SAGE_DK, name=FP))
    put(ws, "A7", "حیاط پشتی ویلای مسکونی · قائم‌شهر، مازندران   |   "
                  "Residential Backyard · Qaemshahr, Mazandaran   |   ≈ 495 m²",
        "band_meta", merge=f"A7:{last}7", fill=sol(CREAM),
        font=fnt(9.5, color=INK_2, name=FP))
    paint(ws, f"A8:{last}8", "rule_gold")

    # --- document control
    put(ws, "A10", "شناسنامه سند  /  DOCUMENT CONTROL", "mini", merge=f"A10:{last}10",
        align_=lambda v: align(v, h="right", wrap=False))
    control = [
        ("پروژه / Project", PROJECT["title_fa"] + " — " + PROJECT["title_en"],
         "کارفرما / Client", PROJECT["client"]),
        ("مکان / Location", PROJECT["place_fa"] + " / " + PROJECT["place_en"],
         "مساحت سایت / Site Area", PROJECT["area"]),
        ("شماره سند / Doc No.", f"ME-LS-00  ·  Rev {REV}",
         "تاریخ / Date", DOC_DATE_EN),
        ("وضعیت / Status", "جهت استعلام / Issued for Pricing",
         "واحد پول / Currency", "تومان / Toman"),
        ("تهیه و طراحی / Prepared & Designed by", "MAJID EISAPOUR",
         "سطح خدمات / Deliverable", "Concept + BOQ + Technical Notes"),
    ]
    r = 11
    for lab1, val1, lab2, val2 in control:
        heights(ws, {r: 23})
        put(ws, f"A{r}", lab1, "label", merge=f"A{r}:B{r}")
        put(ws, f"C{r}", val1, "value", merge=f"C{r}:D{r}")
        put(ws, f"E{r}", lab2, "label", merge=f"E{r}:F{r}")
        put(ws, f"G{r}", val2, "value", merge=f"G{r}:H{r}")
        r += 1

    # --- contents (with internal hyperlinks)
    r += 0
    put(ws, f"A{r}", "فهرست مطالب  /  CONTENTS", "mini", merge=f"A{r}:{last}{r}",
        align_=lambda v: align(v, h="right", wrap=False))
    r += 1
    for i, (code, title, sheet) in enumerate([("00", "جلد و هویت / Cover & Identity", None)] + CONTENTS):
        heights(ws, {r: 20})
        role = "contents" if i % 2 == 0 else "contents_a"
        put(ws, f"A{r}", code, "code_cell")
        put(ws, f"B{r}", title, role, merge=f"B{r}:{last}{r}")
        if sheet:
            ws[f"B{r}"].hyperlink = Hyperlink(ref=f"B{r}", location=f"'{sheet}'!A1")
            ws[f"B{r}"].font = fnt(9.5, color=SAGE_DK, name=FP, italic=False)
        r += 1

    # --- design statement
    r += 1
    put(ws, f"A{r}", "بیانیه طراحی  /  DESIGN STATEMENT", "mini",
        merge=f"A{r}:{last}{r}", align_=lambda v: align(v, h="right", wrap=False))
    r += 1
    heights(ws, {r: 62})
    put(ws, f"A{r}",
        "«فضا فقط ساخته نمی‌شود؛ طراحی می‌شود، تجربه می‌شود و به بخشی از زندگی تبدیل می‌شود.»\n"
        "“Space is not merely built; it is designed, experienced, and made part of everyday life.”",
        "quote", merge=f"A{r}:{last}{r}")
    paint(ws, f"A{r}:{last}{r}", "quote")
    r += 1
    heights(ws, {r: 6})
    paint(ws, f"A{r}:{last}{r}", "spacer")
    r += 1
    foot_strip(ws, r, n,
               f"MAJID EISAPOUR  ·  DESIGN & SPATIAL PRACTICE   |   DOC ME-LS-00  ·  REV {REV}  ·  1405   "
               f"|   CONFIDENTIAL — PREPARED FOR CLIENT REVIEW")

    finish(ws, n, tab=INK)
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToHeight = 1
    ws.print_area = f"A1:{last}{r}"
    ws.oddHeader.left.text = ""
    ws.oddHeader.right.text = ""
    ws.oddFooter.left.text = ""
    ws.oddFooter.right.text = ""
    return ws


def build_info(wb):
    ws = wb.create_sheet("01 | Project Info")
    n = 8
    last = get_column_letter(n)
    widths(ws, {c: 17.5 for c in "ABCDEFGH"})
    band(ws, n, "ME-LS-01", "PROJECT INFORMATION", "مشخصات پروژه",
         "A concise bilingual project identity sheet  |  شناسنامه پروژه", meta2=f"REV {REV} · {DOC_DATE_EN}")
    colheads(ws, 5, ["عنوان / Title", "مقدار / Value"] * 4, height=26)

    fields = [
        ("کارفرما / Client", PROJECT["client"], "مکان / Location", "قائم‌شهر، مازندران / Qaemshahr, Mazandaran",
         "مساحت سایت / Site Area", PROJECT["area"], "تاریخ / Date", "1405 / 2026"),
        ("نوع پروژه / Project Type", "طراحی منظر حیاط ویلایی / Residential Backyard Landscape",
         "زبان طراحی / Design Language", "Contemporary • Natural • Architectural",
         "مقیاس مبنا / Working Scale", "1:100 Concept", "نسخه / Revision", REV),
        ("رویکرد / Approach", "Outdoor Room / اتاقِ بیرونی",
         "برنامه / Program", "Kitchen Garden + Social Space + Pavilion",
         "سطح خدمات / Deliverable", "Concept + BOQ + Technical Notes",
         "تهیه‌کننده / Prepared by", "MAJID EISAPOUR"),
        ("واحد پول / Currency", "تومان / Toman",
         "وضعیت مدرک / Document Status", "جهت استعلام / Issued for Pricing",
         "شماره سند / Doc No.", f"ME-LS-01 · Rev {REV}",
         "مبنای قیمت / Pricing Basis", "برآورد تفصیلی — بدون مالیات بر ارزش افزوده"),
    ]
    r = 6
    for row in fields:
        heights(ws, {r: 40})
        for i in range(0, len(row), 2):
            put(ws, f"{get_column_letter(i + 1)}{r}", row[i], "label")
            put(ws, f"{get_column_letter(i + 2)}{r}", row[i + 1], "value")
        r += 1

    r += 1
    put(ws, f"A{r}", "هویت طراحی / Design Identity", "mini", merge=f"A{r}:{last}{r}",
        align_=lambda v: align(v, h="right", wrap=False))
    r += 1
    heights(ws, {r: 108})
    put(ws, f"A{r}",
        "Design Identity / هویت طراحی\n"
        "A calm, contemporary garden organized as an outdoor room: productive planting to the north, a generous "
        "green heart at the center, and an architectural social space toward the south-west. The service edge "
        "remains intact and is treated as an existing condition rather than a redesign scope.\n\n"
        "باغی معاصر و آرام که همچون یک اتاقِ بیرونی سازمان‌دهی شده است: باغ سبزیِ کاربردی در شمال، قلب سبز و باز "
        "در مرکز، و فضای اجتماعیِ معماری‌شده در جنوب‌غرب. لبه خدماتی موجود دست‌نخورده باقی می‌ماند و صرفاً "
        "به‌عنوان یک وضعیت موجود در نظر گرفته شده است.",
        "body", merge=f"A{r}:{last}{r}")
    paint(ws, f"A{r}:{last}{r}", "body")
    r += 1
    heights(ws, {r: 16})
    put(ws, f"A{r}",
        "تمام مقادیر و مبالغ این دفترچه از مبانی برگه ۰۳ استخراج شده‌اند؛ برای جزئیات به برگه‌های ۰۴ تا ۰۹ مراجعه شود.",
        "note_line", merge=f"A{r}:{last}{r}")

    finish(ws, n, freeze="A6", titles="1:5", tab=INK, footer_left="ME-LS-01")
    return ws


def build_narrative(wb):
    ws = wb.create_sheet("02 | Design Narrative")
    n = 4
    last = get_column_letter(n)
    widths(ws, {"A": 7, "B": 24, "C": 33, "D": 72})
    band(ws, n, "ME-LS-02", "DESIGN NARRATIVE", "روایت و ایده طراحی",
         "Concept, atmosphere and spatial sequence  |  روایت، اتمسفر و توالی فضایی",
         meta2=f"REV {REV}")
    # the narrative table uses A:D only -> merge E:H cells away by painting cream
    colheads(ws, 5, ["ID", "موضوع / Topic", "عنوان / Title", "متن حرفه‌ای / Professional Narrative"],
             height=28)
    r = 6
    for code, topic, title, body in NARRATIVE:
        heights(ws, {r: 74})
        put(ws, f"A{r}", code, "code_cell")
        put(ws, f"B{r}", topic, "row_text" if r % 2 == 0 else "row_text_a")
        put(ws, f"C{r}", title, "row_text" if r % 2 == 0 else "row_text_a",
            font=fnt(9.5, True, color=INK))
        put(ws, f"D{r}", body, "row_text" if r % 2 == 0 else "row_text_a")
        r += 1
    r += 1
    heights(ws, {r: 16})
    put(ws, f"A{r}",
        "توالی فضایی: ورودی از پارکینگ ← آستانه کاشت ← قلب سبز ← پاویون و تراس ← آبنمای خطی",
        "note_line", merge=f"A{r}:{last}{r}")
    finish(ws, n, freeze="A6", titles="1:5", tab=SAGE, footer_left="ME-LS-02")
    return ws


def build_basis(wb):
    ws = wb.create_sheet("03 | Design Basis")
    n = 8
    widths(ws, {"A": 9, "B": 25, "C": 30, "D": 17, "E": 9, "F": 10, "G": 17, "H": 40})
    band(ws, n, "ME-LS-03", "DESIGN BASIS & SCOPE", "مبنای طراحی و دامنه کار",
         "What is included, what remains existing, and the assumptions behind the BOQ  |  مبنای محاسبات",
         meta2=f"REV {REV}")
    colheads(ws, 5, ["کد / Code", "موضوع / Topic", "مبنای طراحی / Basis", "وضعیت / Status",
                     "واحد / Unit", "مقدار / Qty", "یادداشت / Note", "توضیح / Description"],
             height=32)
    r = 6
    for code, topic, basis, status, unit, qty, note, desc in BASIS:
        heights(ws, {r: 30})
        alt = "_a" if r % 2 else ""
        put(ws, f"A{r}", code, "row_code" + alt)
        put(ws, f"B{r}", topic, "row_text" + alt)
        put(ws, f"C{r}", basis, "row_text" + alt)
        put(ws, f"D{r}", status, "row_text" + alt,
            font=fnt(8.5, True, color=SAGE_DK if "Design" in status or "Basis" in status or "Control" in status else TERRA))
        put(ws, f"E{r}", unit, "row_num" + alt)
        put(ws, f"F{r}", qty, "row_num" + alt,
            fmt=QTY if isinstance(qty, (int, float)) else None)
        put(ws, f"G{r}", note, "row_num" + alt,
            font=fnt(8.5, color=INK_2, name=FB))
        put(ws, f"H{r}", desc, "row_note" + alt)
        r += 1
    r += 1
    heights(ws, {r: 30})
    put(ws, f"A{r}",
        "جمع‌بندی / Reconciliation:  مقادیر جدول BOQ (برگه ۰۴) مستقیماً از همین مبانی استخراج شده‌اند؛ "
        "هر تغییر در مبانی، بازنگری مقادیر و مبالغ را الزامی می‌کند.  "
        "اقلام بدون مقدار (—) در دامنه این دفترچه نیستند.",
        "note_line", merge=f"A{r}:{get_column_letter(n)}{r}")
    paint(ws, f"A{r}:{get_column_letter(n)}{r}", "note_line")
    finish(ws, n, freeze="A6", titles="1:5", tab=SAGE, footer_left="ME-LS-03")
    return ws


def build_boq(wb):
    ws = wb.create_sheet("04 | BOQ | برآورد")
    n = 8
    last = get_column_letter(n)
    widths(ws, {"A": 9.5, "B": 26, "C": 44, "D": 11, "E": 10, "F": 17, "G": 19, "H": 46})
    band(ws, n, "ME-LS-04", "BILL OF QUANTITIES — BOQ", "جدول مقادیر و برآورد هزینه",
         "Formula-driven technical estimate  ·  تمام مبالغ به تومان و بدون مالیات بر ارزش افزوده  |  برآورد فرمول‌محور",
         meta2=f"REV {REV}")
    colheads(ws, 5, ["کد / Code", "بخش و آیتم / Item", "شرح فنی و ریزجزئیات / Technical Description",
                     "واحد / Unit", "مقدار / Qty", "قیمت واحد / Unit Rate (Toman)",
                     "مبلغ کل / Total (Toman)", "مبنای محاسبه، ارجاعات و نکات اجرا / Basis & Notes"],
             height=34)

    r = 6
    section_cells = {}          # code -> list of value cells (items only)
    section_subtotal_rows = {}  # code -> subtotal row (if any)
    for sec_no, sec_title, items in BOQ_SECTIONS:
        heights(ws, {r: 20})
        put(ws, f"A{r}", f"{sec_no}", "secbar", h="center")
        put(ws, f"B{r}", sec_title, "secbar", merge=f"B{r}:{last}{r}")
        paint(ws, f"A{r}:{last}{r}", "secbar", border_role="secbar")
        r += 1
        cells = []
        for code, name, desc, unit, qty, rate, note in items:
            heights(ws, {r: 32})
            alt = "_a" if r % 2 else ""
            put(ws, f"A{r}", code, "row_code" + alt)
            put(ws, f"B{r}", name, "row_text" + alt, font=fnt(9, True, color=INK))
            put(ws, f"C{r}", desc, "row_text" + alt)
            put(ws, f"D{r}", unit, "row_num" + alt)
            put(ws, f"E{r}", qty, "row_num" + alt, fmt=QTY)
            put(ws, f"F{r}", rate, "row_money" + alt, fmt=MONEY)
            put(ws, f"G{r}", f"=E{r}*F{r}", "row_money" + alt, fmt=MONEY)
            put(ws, f"H{r}", note, "row_note" + alt)
            cells.append(f"G{r}")
            r += 1
        section_cells[sec_no] = cells
        if len(items) > 1:
            heights(ws, {r: 20})
            put(ws, f"A{r}", "", "subtotal")
            put(ws, f"B{r}", f"جمع بخش {sec_no} / Section {sec_no} Subtotal", "subtotal",
                merge=f"B{r}:F{r}")
            paint(ws, f"A{r}:F{r}", "subtotal")
            put(ws, f"G{r}", f"=SUM({cells[0]}:{cells[-1]})", "subtotal_v", fmt=MONEY)
            put(ws, f"H{r}", f'=IF(ABS(G{r}-SUM({",".join(cells)}))<1,"OK","CHECK")',
                "subtotal", font=fnt(8.5, True, color=GOOD, name=FB))
            section_subtotal_rows[sec_no] = r
            r += 1

    # ---- summary block
    item_refs = [c for cells in section_cells.values() for c in cells]
    direct_expr = "=" + "+".join(item_refs)
    r += 1
    heights(ws, {r: 22})
    put(ws, f"A{r}", "جمع مستقیم / Direct Cost", "grand", merge=f"A{r}:F{r}")
    paint(ws, f"A{r}:F{r}", "grand")
    put(ws, f"G{r}", direct_expr, "grand_v", fmt=MONEY)
    put(ws, f"H{r}", "جمع همه ردیف‌های قیمت‌دار / Sum of all priced items", "grand_v",
        font=fnt(8.5, color=CREAM, name=FP))
    direct_row = r
    r += 1
    heights(ws, {r: 22})
    put(ws, f"A{r}", "ذخیره ریسک و پرت / Risk & Waste Reserve", "grand_alt", merge=f"A{r}:F{r}")
    paint(ws, f"A{r}:F{r}", "grand_alt")
    put(ws, f"G{r}", f"=G{direct_row}*10%", "grand_alt_v", fmt=MONEY)
    put(ws, f"H{r}", "نرخ ذخیره در برگه 09 (سلول ورودی) قابل تنظیم است", "grand_alt_v",
        font=fnt(8.5, color=INK_2, name=FP))
    reserve_row = r
    r += 1
    heights(ws, {r: 24})
    put(ws, f"A{r}", "برآورد کل / Estimated Total", "grand", merge=f"A{r}:F{r}")
    paint(ws, f"A{r}:F{r}", "grand")
    put(ws, f"G{r}", f"=G{direct_row}+G{reserve_row}", "grand_v", fmt=MONEY)
    put(ws, f"H{r}", "شامل ذخیره ریسک و پرت / Includes risk & waste reserve", "grand_v",
        font=fnt(8.5, color=CREAM, name=FP))
    total_row = r
    r += 1
    heights(ws, {r: 30})
    put(ws, f"A{r}",
        "توجه: واحد «مقطوع (LS)» یعنی Lump Sum.  ارجاعات «→» به کدهای سایر برگه‌ها اشاره دارند "
        "(B: مبانی، M: متریال، P: کاشت، L/I: نور و آبیاری، TN: یادداشت‌های اجرایی).",
        "note_line", merge=f"A{r}:{last}{r}")
    paint(ws, f"A{r}:{last}{r}", "note_line")
    finish(ws, n, freeze="A6", titles="1:5", tab=TERRA, footer_left="ME-LS-04")
    return ws, dict(section_cells=section_cells, section_subtotal_rows=section_subtotal_rows,
                    direct_row=direct_row, reserve_row=reserve_row, total_row=total_row)


def build_materials(wb):
    ws = wb.create_sheet("05 | Material Schedule")
    n = 7
    last = get_column_letter(n)
    widths(ws, {"A": 9.5, "B": 30, "C": 26, "D": 30, "E": 10, "F": 11, "G": 52})
    band(ws, n, "ME-LS-05", "MATERIAL SCHEDULE", "جدول متریال و مشخصات",
         "Material language, application and execution intent  |  زبان متریال، کاربرد و قصد اجرایی",
         meta2=f"REV {REV}")
    colheads(ws, 5, ["کد / Code", "متریال / Material", "کاربرد / Application",
                     "حس و کیفیت / Character", "واحد / Unit", "مقدار / Qty",
                     "یادداشت اجرایی / Execution Note"], height=30)
    r = 6
    for code, mat, app, char, unit, qty, note in MATERIALS:
        heights(ws, {r: 40})
        alt = "_a" if r % 2 else ""
        put(ws, f"A{r}", code, "row_code" + alt)
        put(ws, f"B{r}", mat, "row_text" + alt, font=fnt(9, True, color=INK))
        put(ws, f"C{r}", app, "row_text" + alt)
        put(ws, f"D{r}", char, "row_text" + alt, font=fnt(8.5, italic=True, color=SAGE_DK))
        put(ws, f"E{r}", unit, "row_num" + alt)
        put(ws, f"F{r}", qty, "row_num" + alt)
        put(ws, f"G{r}", note, "row_note" + alt)
        r += 1
    r += 1
    heights(ws, {r: 26})
    put(ws, f"A{r}",
        "مقادیر این برگه با برگه ۰۴ (BOQ) و برگه ۰۳ (مبانی) هم‌خوان شده‌اند؛ جمع سنگ 45 + 22 = 67 m² و جمع خاک 28 + 8 = 36 m³.",
        "note_line", merge=f"A{r}:{last}{r}")
    paint(ws, f"A{r}:{last}{r}", "note_line")
    finish(ws, n, freeze="A6", titles="1:5", tab=BRASS, footer_left="ME-LS-05")
    return ws


def build_planting(wb):
    ws = wb.create_sheet("06 | Planting Schedule")
    n = 7
    last = get_column_letter(n)
    widths(ws, {"A": 9.5, "B": 28, "C": 30, "D": 20, "E": 10, "F": 11, "G": 52})
    band(ws, n, "ME-LS-06", "PLANTING SCHEDULE", "پالت کاشت و برنامه سبز",
         "A layered, seasonal and maintainable planting strategy  |  پالت کاشت لایه‌ای، فصلی و قابل نگهداری",
         meta2=f"REV {REV}")
    colheads(ws, 5, ["کد / Code", "گروه / Group", "نقش / Role", "مکان / Zone",
                     "واحد / Unit", "مقدار / Qty", "راهنمای انتخاب / Selection Guide"], height=30)
    r = 6
    for code, grp, role, zone, unit, qty, guide in PLANTING:
        heights(ws, {r: 38})
        alt = "_a" if r % 2 else ""
        put(ws, f"A{r}", code, "row_code" + alt)
        put(ws, f"B{r}", grp, "row_text" + alt, font=fnt(9, True, color=INK))
        put(ws, f"C{r}", role, "row_text" + alt)
        put(ws, f"D{r}", zone, "row_num" + alt)
        put(ws, f"E{r}", unit, "row_num" + alt)
        put(ws, f"F{r}", qty, "row_num" + alt)
        put(ws, f"G{r}", guide, "row_note" + alt)
        r += 1
    r += 1
    heights(ws, {r: 26})
    put(ws, f"A{r}",
        "باغ سبزی: زون 120 m² شامل 90 m² بستر کاشت و 30 m² مسیر و دسترسی  ·  "
        "پرچین 25 متر طول در ردیف 08-04 برآورد شده است.",
        "note_line", merge=f"A{r}:{last}{r}")
    paint(ws, f"A{r}:{last}{r}", "note_line")
    finish(ws, n, freeze="A6", titles="1:5", tab=BRASS, footer_left="ME-LS-06")
    return ws


def build_lighting(wb):
    ws = wb.create_sheet("07 | Lighting & Irrigation")
    n = 7
    last = get_column_letter(n)
    widths(ws, {"A": 9.5, "B": 15, "C": 27, "D": 10, "E": 10, "F": 34, "G": 50})
    band(ws, n, "ME-LS-07", "LIGHTING & IRRIGATION", "نورپردازی و آبیاری",
         "Invisible infrastructure, visible atmosphere  |  زیرساخت نامرئی، اتمسفرِ محسوس",
         meta2=f"REV {REV}")
    colheads(ws, 5, ["کد / Code", "سیستم / System", "عنصر / Element", "واحد / Unit",
                     "مقدار / Qty", "اصل طراحی / Design Principle", "نکته فنی / Technical Note"],
             height=30)
    r = 6
    for code, sysname, elem, unit, qty, princ, note in LIGHTING_IRRIGATION:
        heights(ws, {r: 36})
        alt = "_a" if r % 2 else ""
        put(ws, f"A{r}", code, "row_code" + alt)
        put(ws, f"B{r}", sysname, "row_num" + alt,
            font=fnt(9, True, color=INK_2, name=FB))
        put(ws, f"C{r}", elem, "row_text" + alt)
        put(ws, f"D{r}", unit, "row_num" + alt)
        put(ws, f"E{r}", qty, "row_num" + alt)
        put(ws, f"F{r}", princ, "row_text" + alt, font=fnt(8.5, color=SAGE_DK))
        put(ws, f"G{r}", note, "row_note" + alt)
        r += 1
    r += 1
    heights(ws, {r: 26})
    put(ws, f"A{r}",
        "شبکه اصلی آبیاری (120 lm) و شیرهای برقی (6 عدد) در ردیف 07-01 برآورد لحاظ شده‌اند؛ "
        "خط قطره‌ای 350 lm در ردیف 07-02.",
        "note_line", merge=f"A{r}:{last}{r}")
    paint(ws, f"A{r}:{last}{r}", "note_line")
    finish(ws, n, freeze="A6", titles="1:5", tab=BRASS, footer_left="ME-LS-07")
    return ws


def build_notes(wb):
    ws = wb.create_sheet("08 | Technical Notes")
    n = 6
    last = get_column_letter(n)
    widths(ws, {"A": 10, "B": 30, "C": 46, "D": 18, "E": 18, "F": 18})
    band(ws, n, "ME-LS-08", "TECHNICAL NOTES & QA", "یادداشت‌های اجرایی و کنترل کیفیت",
         "Clear instructions for tender, execution and handover  |  دستور کار روشن برای استعلام، اجرا و تحویل",
         meta2=f"REV {REV}")
    colheads(ws, 5, ["کد / Code", "موضوع / Topic", "دستورالعمل / Instruction", "", "", ""],
             height=28)
    ws.merge_cells(f"C5:{last}5")
    ws["C5"] = "دستورالعمل / Instruction"
    ws["C5"].font = fnt(9.5, True, color=WHITE)
    ws["C5"].fill = sol(INK_2)
    ws["C5"].alignment = align("دستورالعمل", h="center")
    for c in ws[f"C5:{last}5"][0]:
        c.fill = sol(INK_2)
        c.border = bd(color=LINE)
    r = 6
    for code, topic, text in NOTES:
        heights(ws, {r: 42})
        alt = "_a" if r % 2 else ""
        put(ws, f"A{r}", code, "row_code" + alt)
        put(ws, f"B{r}", topic, "row_text" + alt, font=fnt(9, True, color=INK))
        put(ws, f"C{r}", text, "row_text" + alt, merge=f"C{r}:{last}{r}")
        paint(ws, f"C{r}:{last}{r}", "row_text" + alt)
        put(ws, f"C{r}", text, "row_text" + alt)
        r += 1
    r += 1
    heights(ws, {r: 26})
    put(ws, f"A{r}",
        "این یادداشت‌ها مکمل برگه ۰۴ هستند و در زمان استعلام قیمت و اجرا باید به پیمانکار ابلاغ شوند.",
        "note_line", merge=f"A{r}:{last}{r}")
    paint(ws, f"A{r}:{last}{r}", "note_line")
    finish(ws, n, freeze="A6", titles="1:5", tab=SAGE, footer_left="ME-LS-08")
    return ws


def build_cost(wb, boq):
    ws = wb.create_sheet("09 | Cost Summary")
    n = 8
    last = get_column_letter(n)
    widths(ws, {"A": 36, "B": 20, "C": 12, "D": 13, "E": 3, "F": 16, "G": 16, "H": 16})
    band(ws, n, "ME-LS-09", "COST SUMMARY & CONTROL", "خلاصه هزینه و کنترل برآورد",
         "A decision-ready financial snapshot  ·  مبالغ به تومان  |  نمای مالی خوانا و تصمیم‌پذیر",
         meta2=f"REV {REV}")
    colheads(ws, 5, ["بخش / Section", "جمع / Total (Toman)", "سهم / Share", "کنترل / Check",
                     "", "", "", ""], height=28)
    for col in "EFGH":
        ws[f"{col}5"].fill = sol(CREAM)
        ws[f"{col}5"].border = NO_BORDER

    BOQ = "'04 | BOQ | برآورد'"
    sc = boq["section_cells"]

    def sec_sum(codes):
        refs = [c for code in codes for c in sc[code]]
        return f"=SUM({BOQ}!{refs[0]}:{refs[-1]})" if len(refs) > 1 else f"={BOQ}!{refs[0]}"

    groups = [
        ("01–02", "آماده‌سازی و زیرسازی / Preparation & Subbase", ["01", "02"]),
        ("03", "کف‌پوش و محوطه سخت / Paving & Hardscape", ["03"]),
        ("04", "پاویون / Pavilion", ["04"]),
        ("05–06", "تراس و آبنما / Terrace & Water", ["05", "06"]),
        ("07", "آبیاری / Irrigation", ["07"]),
        ("08", "کاشت و چمن / Planting & Lawn", ["08"]),
        ("09", "باغ سبزی / Kitchen Garden", ["09"]),
        ("10", "روشنایی / Lighting", ["10"]),
        ("11", "کنترل کیفیت و تحویل / QA & Handover", ["11"]),
    ]
    r = 6
    first_row = r
    for code, label, codes in groups:
        heights(ws, {r: 26})
        alt = "_a" if r % 2 else ""
        put(ws, f"A{r}", f"{code}  |  {label}", "row_text" + alt)
        put(ws, f"B{r}", sec_sum(codes), "row_money" + alt, fmt=MONEY)
        put(ws, f"C{r}", f"=B{r}/$B${first_row + len(groups)}", "row_num" + alt, fmt=PCT)
        put(ws, f"D{r}", f'=IF(ABS(B{r}-{sec_sum(codes)[1:]})<1,"OK","CHECK")',
            "row_num" + alt, font=fnt(8.5, True, color=GOOD, name=FB))
        r += 1
    last_group_row = r - 1
    direct_row = r
    heights(ws, {r: 24})
    put(ws, f"A{r}", "جمع مستقیم / Direct Cost", "grand_alt")
    paint(ws, f"A{r}:D{r}", "grand_alt")
    put(ws, f"B{r}", f"=SUM(B{first_row}:B{last_group_row})", "grand_alt_v", fmt=MONEY)

    r += 1
    res_row = r
    heights(ws, {r: 24})
    put(ws, f"A{r}", "ذخیره ریسک و پرت / Risk & Waste Reserve", "grand_alt")
    paint(ws, f"A{r}:D{r}", "grand_alt")
    put(ws, f"B{r}", f"=B{direct_row}*C{r}", "grand_alt_v", fmt=MONEY)
    put(ws, f"C{r}", 0.10, "input", fmt='0%')
    put(ws, f"D{r}", "ورودی / editable", "row_num", font=fnt(8, italic=True, color=AMBER_TX, name=FB))

    r += 1
    vat_row = r
    heights(ws, {r: 24})
    put(ws, f"A{r}", "مالیات بر ارزش افزوده / VAT (در صورت اعمال)", "grand_alt")
    paint(ws, f"A{r}:D{r}", "grand_alt")
    put(ws, f"B{r}", f"=B{direct_row}*C{r}", "grand_alt_v", fmt=MONEY)
    put(ws, f"C{r}", 0.0, "input", fmt='0%')
    put(ws, f"D{r}", "پیش‌فرض 0%", "row_num", font=fnt(8, italic=True, color=AMBER_TX, name=FB))

    r += 1
    total_row = r
    heights(ws, {r: 30})
    put(ws, f"A{r}", "برآورد کل / Estimated Total", "grand")
    paint(ws, f"A{r}:D{r}", "grand")
    put(ws, f"B{r}", f"=B{direct_row}+B{res_row}+B{vat_row}", "grand_v", fmt=MONEY,
        font=fnt(11.5, True, color=WHITE, name=FB))
    put(ws, f"C{r}", f"=B{total_row}/B{direct_row}", "grand_v", fmt=PCT)
    put(ws, f"D{r}", "کنترل / Control", "grand_v", font=fnt(8, color=CREAM, name=FP))

    # ---- KPI block
    r += 2
    put(ws, f"A{r}", "شاخص‌های کلیدی  /  KEY INDICATORS", "mini",
        merge=f"A{r}:D{r}", align_=lambda v: align(v, h="right", wrap=False))
    kpi_start = r + 1
    kpis = [
        ("هزینه بر مترمربع سایت / Cost per site m²",
         f"=B{total_row}/495", MONEY),
        ("سهم محوطه سخت، تراس و آب / Paving, terrace & water",
         f"=(B{first_row + 1}+B{first_row + 3})/B{total_row}", PCT),
        ("سهم کاشت، چمن و باغ سبزی / Planting, lawn & kitchen garden",
         f"=(B{first_row + 5}+B{first_row + 6})/B{total_row}", PCT),
        ("سهم پاویون / Pavilion share",
         f"=B{first_row + 2}/B{total_row}", PCT),
    ]
    r = kpi_start
    for label, formula, fmt in kpis:
        heights(ws, {r: 24})
        put(ws, f"A{r}", label, "row_text" + ("_a" if r % 2 else ""))
        put(ws, f"B{r}", formula, "kpi_val", fmt=fmt)
        put(ws, f"C{r}", "", "row_text")
        put(ws, f"D{r}", "", "row_text")
        r += 1

    r += 1
    heights(ws, {r: 34})
    put(ws, f"A{r}",
        "کنترل برآورد:  برگه ۰۴ مبنا است و این برگه فقط جمع‌بندی می‌کند.  ستون «کنترل» مغایرت جمع گروه با ردیف‌های BOQ را "
        "هشدار می‌دهد و ستون «سهم» بر پایه برآورد کل محاسبه می‌شود.",
        "note_line", merge=f"A{r}:D{r}")
    paint(ws, f"A{r}:D{r}", "note_line")

    # ---- conditional formatting (data bars) + chart
    ws.conditional_formatting.add(
        f"C{first_row}:C{last_group_row}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=0.4,
                    color=SAGE, showValue=True, minLength=None, maxLength=None))

    chart = BarChart()
    chart.type = "bar"
    chart.style = 2
    chart.title = "سهم هزینه بخش‌ها  |  Section Cost Share (Toman)"
    chart.y_axis.title = "Toman"
    chart.x_axis.title = "Section"
    chart.gapWidth = 55
    chart.height = 10.5
    chart.width = 15
    data = Reference(ws, min_col=2, min_row=first_row, max_row=last_group_row)
    cats = Reference(ws, min_col=1, min_row=first_row, max_row=last_group_row)
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)
    chart.series[0].cat = AxDataSource(
        strRef=StrRef(f"'09 | Cost Summary'!$A${first_row}:$A${last_group_row}"))
    chart.legend = None
    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.numFmt = '#,##0'
    s = chart.series[0]
    s.graphicalProperties.solidFill = TERRA
    s.graphicalProperties.line.solidFill = TERRA
    chart.x_axis.scaling.orientation = "maxMin"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.numFmt = '#,##0'
    chart.y_axis.majorGridlines.spPr = GraphicalProperties()
    chart.y_axis.majorGridlines.spPr.line = LineProperties(solidFill=LINE_LT)
    ws.add_chart(chart, "F6")

    finish(ws, n, freeze="A6", titles="1:5", tab=TERRA, footer_left="ME-LS-09")
    ws.print_area = f"A1:{last}{r}"
    return dict(ws=ws, direct_row=direct_row, res_row=res_row, vat_row=vat_row,
                total_row=total_row)


def build_revision(wb):
    ws = wb.create_sheet("10 | Revision Log")
    n = 6
    last = get_column_letter(n)
    widths(ws, {"A": 10, "B": 20, "C": 62, "D": 20, "E": 16, "F": 16})
    band(ws, n, "ME-LS-10", "REVISION LOG & DOCUMENT CONTROL", "تاریخچه نسخه‌ها و کنترل مدارک",
         "Keep every issue traceable  |  کنترل نسخه و گردش مستندات", meta2=f"REV {REV}")
    colheads(ws, 5, ["نسخه / Rev.", "تاریخ / Date", "شرح تغییر / Description",
                     "تهیه / Prepared", "بررسی / Checked", "تأیید / Approved"], height=28)
    r = 6
    for rev, date, desc, prep, chk, appr in REVISIONS:
        heights(ws, {r: 46 if rev == "02" else 34})
        alt = "_a" if r % 2 else ""
        put(ws, f"A{r}", rev, "row_code" + alt)
        put(ws, f"B{r}", date, "row_num" + alt)
        put(ws, f"C{r}", desc, "row_text" + alt)
        put(ws, f"D{r}", prep, "row_num" + alt, font=fnt(8.5, color=INK_2, name=FB))
        put(ws, f"E{r}", chk, "row_num" + alt)
        put(ws, f"F{r}", appr, "row_num" + alt)
        r += 1
    r += 1
    heights(ws, {r: 30})
    put(ws, f"A{r}",
        "این دفترچه به‌صورت اسکریپت‌محور بازتولید می‌شود (tools/build_workbook.py)؛ هر اصلاح بعدی با نسخه جدید "
        "ثبت و قابل ردیابی است.",
        "note_line", merge=f"A{r}:{last}{r}")
    paint(ws, f"A{r}:{last}{r}", "note_line")
    finish(ws, n, freeze="A6", titles="1:5", tab=INK_2, footer_left="ME-LS-10")
    return ws


# ---------------------------------------------------------------------------
def main():
    wb = Workbook()
    wb.remove(wb.active)

    build_cover(wb)
    build_info(wb)
    build_narrative(wb)
    build_basis(wb)
    boq_ws, boq = build_boq(wb)
    build_materials(wb)
    build_planting(wb)
    build_lighting(wb)
    build_notes(wb)
    cost = build_cost(wb, boq)
    build_revision(wb)

    # single source of truth for the reserve rate: BOQ reads sheet 09 input cell
    boq_ws[f"G{boq['reserve_row']}"] = (
        f"=G{boq['direct_row']}*'09 | Cost Summary'!$C${cost['res_row']}"
    )

    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    props = wb.properties
    props.creator = "MAJID EISAPOUR"
    props.lastModifiedBy = "MAJID EISAPOUR"
    props.title = "MAJID EISAPOUR | Landscape Design & Technical BOQ | Qaemshahr"
    props.subject = "Bilingual premium landscape BOQ and project documentation — Rev 02"
    props.description = ("Prepared & Designed by MAJID EISAPOUR — Design & Spatial Practice. "
                         "Rebuilt revision 02: graphic redesign, quantity reconciliation, formula and print fixes.")
    props.keywords = "landscape; BOQ; Qaemshahr; landscape architecture; برآورد; طراحی منظر"
    props.category = "Landscape Architecture — Technical Documentation"
    props.created = _dt.datetime(2026, 9, 26, 18, 35, 51)

    wb.save(OUT)
    print("✅ saved:", OUT)
    for ws in wb.worksheets:
        print(f"   · {ws.title:<28} rows={ws.max_row:<3} cols={ws.max_column:<2} "
              f"print={ws.print_area}")


if __name__ == "__main__":
    main()
