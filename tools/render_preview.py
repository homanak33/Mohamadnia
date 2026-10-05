#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MAJID EISAPOUR — Landscape Design & Technical BOQ
preview renderer:  XLSX  ➜  self-contained HTML (نمایش وب دفترچه)

این ابزار فایل اکسل را می‌خواند، فرمول‌ها را ارزیابی می‌کند و هر برگه را
دقیقاً با همان گرافیک (رنگ، فونت، ادغام سلول‌ها، ابعاد سطر/ستون و RTL)
به یک صفحه HTML تبدیل می‌کند تا دفترچه بدون اکسل هم قابل مشاهده و ارائه باشد.

Run:  python3 tools/render_preview.py
"""
from __future__ import annotations

import html
import os
import unicodedata
import re

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(REPO, "MAJID_EISAPOUR_QAEMSHAHR_LANDSCAPE_BOQ_PREMIUM.xlsx")
OUT_DIR = os.path.join(REPO, "preview")
OUT_HTML = os.path.join(OUT_DIR, "index.html")

# --------------------------------------------------------------------------
# 1 | formula engine  (enough for the BOQ / cost sheet formulas)
# --------------------------------------------------------------------------
CELL_RE = re.compile(r"(?:'([^']+)'!|([A-Za-z0-9_.&| آ-ی]+)!)?(\$?[A-Z]{1,3}\$?\d+)")


class Engine:
    def __init__(self, wb):
        self.wb = wb
        self.raw = {}
        for ws in wb.worksheets:
            cells = {}
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        cells[c.coordinate] = c.value
            self.raw[ws.title] = cells
        self.memo = {}

    # -- helpers ---------------------------------------------------------
    def _split(self, text):
        out, depth, buf, quote = [], 0, "", False
        for ch in text:
            if ch == '"':
                quote = not quote
            if not quote:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                elif ch == "," and depth == 0:
                    out.append(buf)
                    buf = ""
                    continue
            buf += ch
        out.append(buf)
        return out

    def cell(self, sheet, coord, seen=None):
        key = (sheet, coord)
        if key in self.memo:
            return self.memo[key]
        seen = seen or set()
        if key in seen:
            return 0
        seen = seen | {key}
        val = self.raw.get(sheet, {}).get(coord, 0)
        if isinstance(val, str) and val.startswith("="):
            val = self.eval(val, sheet, seen)
        self.memo[key] = val
        return val

    def _num(self, v):
        if v is None:
            return 0.0
        if isinstance(v, bool):
            return float(v)
        if isinstance(v, (int, float)):
            return float(v)
        try:
            return float(str(v).replace(",", "").replace("٬", ""))
        except ValueError:
            return 0.0

    def _resolve(self, ref, sheet):
        ref = ref.strip().replace("$", "")
        if "!" in ref:
            sh, _, ref = ref.rpartition("!")
            sh = sh.strip().strip("'")
        else:
            sh = sheet
        return self.cell(sh, ref.upper())

    def _range_sum(self, ref, sheet):
        if "!" in ref:
            sh, _, rng = ref.rpartition("!")
            sh = sh.strip().strip("'")
        else:
            sh, rng = sheet, ref
        rng = rng.replace("$", "")
        min_c, min_r, max_c, max_r = range_boundaries(rng)
        total = 0.0
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                total += self._num(self.cell(sh, f"{get_column_letter(c)}{r}"))
        return total

    # -- evaluator -------------------------------------------------------
    def eval(self, expr, sheet, seen=None):
        e = expr[1:] if expr.startswith("=") else expr

        def sum_repl(m):
            total = 0.0
            for part in self._split(m.group(1)):
                part = part.strip()
                if not part:
                    continue
                if ":" in part:
                    total += self._range_sum(part, sheet)
                else:
                    total += self._num(self._resolve(part, sheet))
            return repr(total)

        e = re.sub(r"SUM\(([^()]*)\)", sum_repl, e)
        e = re.sub(r"\bIF\(", "ifelse(", e)
        e = re.sub(r"\bABS\(", "abs(", e)
        e = re.sub(r"(\d+(?:\.\d+)?)%", r"(\1/100)", e)
        e = e.replace("<>", "!=").replace("^", "**")
        e = re.sub(r"(?<![<>=!])=(?!=)", "==", e)

        def ref_repl(m):
            sh = m.group(1)
            coord = m.group(3).replace("$", "")
            if sh:
                return repr(self._num(self.cell(sh.strip().strip("'"), coord.upper())))
            return repr(self._num(self._resolve(coord, sheet)))

        e = CELL_RE.sub(ref_repl, e)
        try:
            return eval(e, {"__builtins__": {}}, {"ifelse": lambda c, a, b: a if c else b,
                                                   "abs": abs})
        except Exception:
            return None

    def display(self, sheet, coord, value):
        if isinstance(value, str) and value.startswith("="):
            return self.eval(value, sheet)
        return value


# --------------------------------------------------------------------------
# 2 | formatting helpers
# --------------------------------------------------------------------------
def fmt_number(value, number_format):
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        return value
    nf = (number_format or "General").lower()
    try:
        num = float(value)
    except (TypeError, ValueError):
        return str(value)
    if "%" in nf:
        dec = 1 if "0.0" in nf else 0
        return f"{num * 100:,.{dec}f}%"
    if nf in ("general", "@"):
        return (f"{num:,.10f}".rstrip("0").rstrip(".") if num % 1 else f"{num:,.0f}")
    dec = 2 if "#,##0.##" in nf else (1 if "0.0" in nf else 0)
    return f"{num:,.{dec}f}"


def has_fa(text) -> bool:
    """first strong directional character decides (Excel General alignment rule)"""
    for ch in str(text or ""):
        bidi = unicodedata.bidirectional(ch)
        if bidi in ("R", "AL"):
            return True
        if bidi == "L":
            return False
    return False


def px_width(col_width):
    return int(round((col_width or 8.43) * 7)) + 5


def px_height(row_height):
    return int(round((row_height or 15) * 1.34))


def css_color(color):
    if color is None:
        return None
    if getattr(color, "type", None) == "rgb" and color.rgb:
        rgb = str(color.rgb)
        if rgb in ("00000000",):
            return None
        return "#" + rgb[-6:]
    if getattr(color, "type", None) == "theme":
        return None
    return None


def css_align(a, text=None):
    """Mirror Excel behaviour: 'general' follows the first strong character."""
    if a is None:
        return {}
    out = {}
    h = a.horizontal
    if h in (None, "general"):
        out["text-align"] = "right" if has_fa(text) else "left"
    else:
        out["text-align"] = h
    if a.vertical:
        out["vertical-align"] = {"center": "middle", "top": "top",
                                 "bottom": "bottom"}.get(a.vertical, "middle")
    if a.wrap_text:
        out["white-space"] = "pre-wrap"
    return out


def font_css(f):
    if f is None:
        return {}
    out = {}
    fam = (f.name or "Tahoma")
    if "Aptos Display" in fam:
        out["font-family"] = "'Aptos Display','Segoe UI',Calibri,Arial,sans-serif"
    elif "Aptos" in fam:
        out["font-family"] = "Aptos,'Segoe UI',Calibri,Arial,sans-serif"
    else:
        out["font-family"] = "Tahoma,'Segoe UI',Arial,sans-serif"
    if f.size:
        out["font-size"] = f"{float(f.size) * 1.30:.1f}px"
    if f.bold:
        out["font-weight"] = "700"
    if f.italic:
        out["font-style"] = "italic"
    col = css_color(f.color)
    if col:
        out["color"] = col
    return out


def border_css(b):
    if b is None:
        return {}
    out = {}
    for side, key in (("left", "border-left"), ("right", "border-right"),
                      ("top", "border-top"), ("bottom", "border-bottom")):
        sd = getattr(b, side, None)
        if sd is None or not sd.style:
            continue
        width = {"thin": "1px", "medium": "2px", "thick": "3px", "dashed": "1px",
                 "dotted": "1px", "double": "3px", "hair": "1px"}.get(sd.style, "1px")
        style = "dashed" if sd.style == "dashed" else ("dotted" if sd.style == "dotted" else "solid")
        color = css_color(sd.color) or "#D8D1C0"
        out[key] = f"{width} {style} {color}"
    return out


def fill_css(f):
    if f is None:
        return {}
    tagname = getattr(f, "tagname", "") or ""
    kind = getattr(f, "fill_type", None)
    if tagname == "gradientFill" or kind == "gradient":
        stops = getattr(f, "stop", None) or []
        if len(stops) >= 2:
            c1 = css_color(stops[0].color) or "#1E2A26"
            c2 = css_color(stops[-1].color) or "#33413B"
            degree = getattr(f, "degree", 0) or 0
            angle = (90 - degree) % 360
            return {"background": f"linear-gradient({angle}deg,{c1},{c2})"}
        return {}
    if kind is None:
        return {}
    col = css_color(f.fgColor)
    return {"background": col} if col else {}


# --------------------------------------------------------------------------
# 3 | renderer
# --------------------------------------------------------------------------
def render_sheet(ws, engine, index, total):
    rows = ws.max_row
    cols = ws.max_column
    merged = {}
    covered = set()
    for rng in ws.merged_cells.ranges:
        min_c, min_r, max_c, max_r = range_boundaries(str(rng))
        merged[(min_r, min_c)] = (max_r - min_r + 1, max_c - min_c + 1)
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                if (r, c) != (min_r, min_c):
                    covered.add((r, c))

    colgroup = "".join(
        f'<col style="width:{px_width(ws.column_dimensions[get_column_letter(c)].width)}px">'
        for c in range(1, cols + 1))

    body_rows = []
    for r in range(1, rows + 1):
        rdim = ws.row_dimensions.get(r)
        cells_html = []
        for c in range(1, cols + 1):
            if (r, c) in covered:
                continue
            cell = ws.cell(row=r, column=c)
            span = merged.get((r, c))
            attrs = []
            if span:
                attrs.append(f'rowspan="{span[0]}"')
                attrs.append(f'colspan="{span[1]}"')
            value = engine.display(ws.title, cell.coordinate, cell.value)
            style = {}
            style.update(font_css(cell.font))
            style.update(fill_css(cell.fill))
            style.update(border_css(cell.border))
            style.update(css_align(cell.alignment, value))
            text = html.escape(fmt_number(value, cell.number_format)) if value is not None else ""
            if style:
                attrs.append("style=\"" + ";".join(f"{k}:{v}" for k, v in style.items()) + "\"")
            cls = "cell"
            if cell.hyperlink is not None and cell.hyperlink.location:
                cls += " link"
            cells_html.append(f'<td class="{cls}" {" ".join(attrs)}>{text}</td>')
        h = f' style="height:{px_height(rdim.height if rdim else None)}px"'
        body_rows.append(f"<tr{h}>" + "".join(cells_html) + "</tr>")

    active = " active" if index == 1 else ""
    return f'''
<section class="sheet{active}" id="sheet-{index}" data-title="{html.escape(ws.title)}">
  <div class="sheetbar"><span class="tabdot"></span>{html.escape(ws.title)}</div>
  <div class="page">
    <table class="grid" dir="rtl">
      <colgroup>{colgroup}</colgroup>
      <tbody>{"".join(body_rows)}</tbody>
    </table>
  </div>
  <div class="pagefoot">
    <span>MAJID EISAPOUR · Design &amp; Spatial Practice</span>
    <span>{html.escape(ws.title)}</span>
    <span>صفحه {index} از {total}</span>
  </div>
</section>'''


CSS = """
:root{
  --ink:#1E2A26; --ink2:#33413B; --sage:#7C8B72; --sagedk:#59684F; --sagelt:#E7EBDF;
  --terra:#B4713F; --brass:#C0A163; --cream:#F8F5EC; --cream2:#F1EBDD; --line:#D8D1C0;
}
*{box-sizing:border-box}
body{margin:0;background:#20261f;color:#1E2A26;
     font-family:Tahoma,'Segoe UI',Arial,sans-serif;direction:rtl}
header.top{position:sticky;top:0;z-index:20;background:linear-gradient(100deg,#16201C,#28332D);
  border-bottom:3px solid var(--terra);padding:10px 18px;display:flex;flex-wrap:wrap;
  gap:14px;align-items:center;justify-content:space-between;box-shadow:0 6px 22px rgba(0,0,0,.35)}
header.top .brand{display:flex;align-items:center;gap:12px}
header.top .mono{width:42px;height:42px;border:1px solid var(--brass);color:var(--brass);
  display:grid;place-items:center;font-family:'Aptos Display',Segoe UI,sans-serif;
  font-weight:800;letter-spacing:.5px;font-size:15px;border-radius:2px}
header.top h1{margin:0;color:#F8F5EC;font-size:15px;font-weight:700;letter-spacing:.2px}
header.top .sub{color:var(--brass);font-size:11px;margin-top:2px}
nav{display:flex;flex-wrap:wrap;gap:6px}
nav button{background:transparent;border:1px solid #4d5c50;color:#dfe5d8;border-radius:2px;
  padding:5px 9px;font-size:11px;cursor:pointer;font-family:inherit;transition:.15s}
nav button:hover{background:#2c3a31;border-color:var(--sage)}
nav button.active{background:var(--terra);border-color:var(--terra);color:#fff}
main{padding:22px 16px 60px;display:flex;flex-direction:column;gap:26px;align-items:center}
section.sheet{display:none;width:100%;max-width:1500px}
section.sheet.active{display:block;animation:fade .25s ease}
@keyframes fade{from{opacity:0;transform:translateY(6px)}to{opacity:1}}
.sheetbar{color:#cfd8c8;font-size:11.5px;padding:0 4px 8px;display:flex;gap:8px;align-items:center}
.tabdot{width:9px;height:9px;background:var(--terra);display:inline-block;border-radius:50%}
.page{background:#fff;border-radius:3px;overflow:auto;box-shadow:0 14px 40px rgba(0,0,0,.45);
  border-top:4px solid var(--ink)}
table.grid{border-collapse:collapse;table-layout:fixed;width:max-content;min-width:100%}
table.grid td{overflow:hidden;word-break:break-word;line-height:1.45;padding:0 4px}
table.grid td.link{text-decoration:underline;text-decoration-color:var(--brass)}
.pagefoot{display:flex;justify-content:space-between;color:#8d9a8b;font-size:10.5px;padding:7px 4px 0}
footer.credit{color:#7d8a7c;font-size:11px;text-align:center;padding:0 0 30px}
footer.credit b{color:#c8d2c2}
@media print{
  body{background:#fff}
  header.top,nav,.sheetbar,.pagefoot,footer.credit{display:none!important}
  section.sheet{display:block!important;page-break-after:always;max-width:none}
  .page{box-shadow:none;border-top:none;overflow:visible}
  table.grid td{font-size:9px}
}
"""

JS = """
const buttons=[...document.querySelectorAll('nav button')];
const sheets=[...document.querySelectorAll('section.sheet')];
function show(i){
  sheets.forEach((s,k)=>s.classList.toggle('active',k===i));
  buttons.forEach((b,k)=>b.classList.toggle('active',k===i));
  window.scrollTo({top:0,behavior:'smooth'});
}
buttons.forEach((b,i)=>b.addEventListener('click',()=>show(i)));
document.addEventListener('keydown',e=>{
  const cur=sheets.findIndex(s=>s.classList.contains('active'));
  if(e.key==='ArrowLeft') show(Math.min(cur+1,sheets.length-1));
  if(e.key==='ArrowRight') show(Math.max(cur-1,0));
});
"""


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    wb = openpyxl.load_workbook(XLSX)
    engine = Engine(wb)

    nav = "".join(
        f'<button class="{"active" if i == 0 else ""}">{html.escape(ws.title)}</button>'
        for i, ws in enumerate(wb.worksheets))

    pages = []
    for i, ws in enumerate(wb.worksheets, start=1):
        pages.append(render_sheet(ws, engine, i, len(wb.worksheets)))

    doc = f"""<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>MAJID EISAPOUR · Landscape Design &amp; Technical BOQ · Qaemshahr</title>
<style>{CSS}</style>
</head>
<body>
<header class="top">
  <div class="brand">
    <div class="mono">ME</div>
    <div>
      <h1>MAJID EISAPOUR — Landscape Design &amp; Technical BOQ</h1>
      <div class="sub">حیاط پشتی ویلای مسکونی · قائم‌شهر، مازندران · Rev 02 · 1405</div>
    </div>
  </div>
  <nav>{"".join(nav)}</nav>
</header>
<main>{"".join(pages)}</main>
<footer class="credit">نمایش وب دفترچه — بازتولید از فایل <b>MAJID_EISAPOUR_QAEMSHAHR_LANDSCAPE_BOQ_PREMIUM.xlsx</b>
 · برای چاپ: Ctrl/Cmd + P</footer>
<script>{JS}</script>
</body>
</html>"""

    with open(OUT_HTML, "w", encoding="utf-8") as fh:
        fh.write(doc)
    print("✅ preview:", OUT_HTML, f"({os.path.getsize(OUT_HTML)/1024:.0f} KB)")


if __name__ == "__main__":
    main()
