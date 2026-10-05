# MAJID EISAPOUR — Landscape Design & Technical BOQ

**Project:** Residential Backyard Villa — Qaemshahr, Mazandaran
**Document:** دفترچه طراحی منظر و برآورد فنی · Rev 02 · 1405/2026

| File | Description |
|---|---|
| `MAJID_EISAPOUR_QAEMSHAHR_LANDSCAPE_BOQ_PREMIUM.xlsx` | دفترچه نهایی (۱۱ برگه، دوزبانه، بازطراحی‌شده) |
| `preview/index.html` | نمایش وب دفترچه با همان گرافیک (بدون اکسل) |
| `AUDIT_REV02_FA.md` | گزارش کامل بازبینی، اصلاحات و تصمیم‌ها (نسخه ۰۲) |
| `tools/build_workbook.py` | منبع واحد ساخت فایل اکسل |
| `tools/render_preview.py` | ساخت پیش‌نمایش وب از روی فایل اکسل |
| `archive/` | نسخه اصلی و دست‌نخورده (نسخه ۰۱) |

## Build

```bash
python3 tools/build_workbook.py     # rebuild the workbook
python3 tools/render_preview.py     # rebuild the web preview
```

> The workbook is generated — always edit `tools/build_workbook.py` (single source of truth)
> instead of editing the `.xlsx` by hand, so every sheet stays consistent.
