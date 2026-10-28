#!/usr/bin/env python3
"""Build Freelancer-Expense-Category-Sample.xlsx — free structural seam of expense-sheet-v1."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import date
from pathlib import Path

OUT = Path(__file__).resolve().parent / "Freelancer-Expense-Category-Sample.xlsx"

H = Font(bold=True, color="FFFFFF")
HF = PatternFill("solid", fgColor="2F4F4F")
IN = PatternFill("solid", fgColor="FFF8DC")
T = Font(bold=True, size=16)
B = Font(bold=True)


def head(ws, row, cols, widths):
    for i, (c, w) in enumerate(zip(cols, widths), 1):
        x = ws.cell(row, i, c)
        x.font = H
        x.fill = HF
        x.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[x.column_letter].width = w
    ws.freeze_panes = ws.cell(row + 1, 1)


def build():
    wb = Workbook()
    s = wb.active
    s.title = "Start Here"
    lines = [
        "Freelancer Expense Category Starter — Free Sample",
        "",
        "What this is",
        "A complete, usable one-month expense log with Schedule C–style categories. Same quality as the paid tracker’s Categories + Expenses core — not a crippled demo.",
        "",
        "What you can do in 10 minutes",
        "1) Skim Categories — rename or add rows if your tax authority uses different labels.",
        "2) Replace EXAMPLE rows on Expense Log (One Month) with your January (or any one month) purchases.",
        "3) Read Month totals — spent, deductible (after business-use %), by category.",
        "",
        "Seam map (what this deliberately does NOT do)",
        "FREE sample includes: Categories list + one-month Expense Log with auto deductible / year / month.",
        "NOT in free sample (paid Freelancer Expense Tracker breadth):",
        "  • Full-year Expenses grid (1000+ rows) across all months",
        "  • Mileage log with rate from Settings",
        "  • Dashboard: tax-year KPIs + category + month rollups for your accountant",
        "  • Settings sheet (currency, tax year, mileage rate)",
        "  • Receipt-link + payment-method columns at full scale",
        "",
        "Gap questions (reply with one word — this is the real demand signal)",
        "Downloads prove the funnel exists. A specific question proves purchase intent. Pick one and reply (email / DM / GitHub issue):",
        "1) SCHEDULEC — Do you want the full Schedule C category set wired to a year-long log + accountant-ready category totals?",
        "2) MILEAGE — Do you want a mileage log that multiplies distance by your authority’s rate automatically?",
        "3) DASHBOARD — Do you want tax-year spent / deductible / by-month rollups without building a second sheet?",
        "If none apply, keep the free sheet — no pitch. If one word fits, that is the paid Freelancer Expense Tracker seam.",
        "",
        "What “missing” looks like (not a vague “buy full pack”)",
        "Free sample solves: categorize and total one month of business spend. Paid adds: Settings + Mileage + full-year Expenses + Dashboard so tax-prep is one file.",
        "",
        "License",
        "Free for personal use. Example data is fictional. Categories follow common US IRS Schedule C line labels (public methodology) — rename for your jurisdiction. Not tax advice. Gap-question CTA added 2026-10-28 (free-expense-sample-v1).",
    ]
    for i, l in enumerate(lines, 1):
        s.cell(i, 1, l)
    s["A1"].font = T
    s.column_dimensions["A"].width = 120

    cats = [
        "Advertising",
        "Car and truck expenses",
        "Commissions and fees",
        "Contract labor",
        "Insurance",
        "Legal and professional services",
        "Office expense",
        "Supplies",
        "Travel",
        "Meals",
        "Utilities",
        "Software & subscriptions",
        "Education & training",
        "Phone & internet",
        "Bank & payment fees",
        "Other expenses",
    ]
    notes = {
        "Meals": "Often only partly deductible; check your rules",
        "Software & subscriptions": "Often falls under Other expenses on Schedule C",
        "Phone & internet": "Use business-use % for mixed personal/work bills",
    }
    ca = wb.create_sheet("Categories")
    head(ca, 1, ["Category", "Notes"], [34, 60])
    for r, c in enumerate(cats, 2):
        ca.cell(r, 1, c).fill = IN
        ca.cell(r, 2, notes.get(c, ""))

    ex = wb.create_sheet("Expense Log (One Month)")
    head(
        ex,
        1,
        ["Date", "Vendor", "Description", "Category", "Amount", "Business-use %", "Deductible", "Year", "Month"],
        [12, 22, 32, 28, 12, 14, 12, 8, 8],
    )
    dv = DataValidation(type="list", formula1="=Categories!$A$2:$A$50", allow_blank=True)
    ex.add_data_validation(dv)

    sample = [
        (date(2026, 1, 5), "EXAMPLE Adobe", "Creative Cloud monthly", "Software & subscriptions", 59.99, 1),
        (date(2026, 1, 8), "EXAMPLE Staples", "Printer paper + ink", "Office expense", 48.20, 1),
        (date(2026, 1, 12), "EXAMPLE Verizon", "Phone bill (work share)", "Phone & internet", 80.00, 0.6),
        (date(2026, 1, 15), "EXAMPLE Meta", "Instagram ads", "Advertising", 120.00, 1),
        (date(2026, 1, 20), "EXAMPLE Uber", "Client meeting travel", "Travel", 34.50, 1),
        (date(2026, 1, 22), "EXAMPLE Cafe", "Client lunch (check meal rules)", "Meals", 42.00, 0.5),
        (date(2026, 1, 28), "EXAMPLE Upwork", "Platform fee", "Commissions and fees", 38.00, 1),
    ]
    n = len(sample)

    for r in range(2, 52):
        for c in (1, 2, 3, 4, 5, 6):
            ex.cell(r, c).fill = IN
        dv.add(f"D{r}")
        ex.cell(r, 1).number_format = "yyyy-mm-dd"
        ex.cell(r, 5).number_format = "#,##0.00"
        ex.cell(r, 6).number_format = "0%"
        ex.cell(r, 7, f'=IF(E{r}="","",ROUND(E{r}*IF(F{r}="",1,F{r}),2))').number_format = "#,##0.00"
        ex.cell(r, 8, f'=IF(A{r}="","",YEAR(A{r}))')
        ex.cell(r, 9, f'=IF(A{r}="","",MONTH(A{r}))')

    for r, row in enumerate(sample, 2):
        for c, v in zip((1, 2, 3, 4, 5, 6), row):
            ex.cell(r, c, v)

    ex.cell(54, 1, "Month totals (auto)").font = B
    ex.cell(55, 1, "Total spent")
    ex.cell(55, 2, f"=SUM(E2:E{1+n})").number_format = "#,##0.00"
    ex.cell(56, 1, "Total deductible")
    ex.cell(56, 2, f"=SUM(G2:G{1+n})").number_format = "#,##0.00"
    ex.cell(58, 1, "By category").font = B
    ex.cell(59, 1, "Category").font = H
    ex.cell(59, 1).fill = HF
    ex.cell(59, 2, "Deductible").font = H
    ex.cell(59, 2).fill = HF
    for i, cname in enumerate(cats):
        rr = 60 + i
        ex.cell(rr, 1, cname)
        ex.cell(rr, 2, f"=SUMIF($D$2:$D${1+n},A{rr},$G$2:$G${1+n})").number_format = "#,##0.00"

    ex.cell(78, 1, "Tip: delete EXAMPLE rows after you paste your own. Yellow = inputs. Deductible / Year / Month are formulas.")
    ex.cell(79, 1, "This sample is one month only. Paid Freelancer Expense Tracker adds Mileage + Dashboard + full-year log.")

    wb.save(OUT)
    print("ok", OUT)
    print("hand-calc expected: spent 422.69, deductible 369.69")


if __name__ == "__main__":
    build()
