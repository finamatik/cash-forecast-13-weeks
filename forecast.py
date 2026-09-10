"""13-week cash flow forecast built from the receivables and payables ledgers plus a recurring schedule.

Input (CSV, as exported from any accounting system):
  sample/ar_open.csv      invoice, customer, issue_date, due_date, amount_aed
  sample/ap_open.csv      bill, supplier, due_date, amount_aed, category
  sample/recurring.csv    name, category, amount_aed, frequency (weekly|monthly), day_or_date, direction (in|out)
  sample/customers.csv    customer, avg_days_late  (payment behaviour from history; drives expected receipt dates)
Output: out/cash-forecast-13w.xlsx with Forecast, Scenarios, Receipts detail, Payments detail, Assumptions, and a chart.
Read only. Deterministic. No AI."""
from __future__ import annotations

import csv
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
INK, IVORY, BRASS, MUTED, LINE = "0B1220", "F6F1E7", "C9A24C", "6B7280", "E4DCC9"
HEAD = PatternFill("solid", fgColor=INK)
BAND = PatternFill("solid", fgColor="FBF8F1")
RED = PatternFill("solid", fgColor="FDE8E8")
thin = Side(style="thin", color=LINE)
BORDER = Border(bottom=thin)


def d(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


def rows(name):
    with open(HERE / "sample" / name, newline="") as f:
        return list(csv.DictReader(f))


def week_index(day: date, start: date) -> int:
    return (day - start).days // 7


def build(start: date, opening_cash: Decimal, weeks: int = 13, slow_days: int = 0, lost_customer: str | None = None, ap_stretch: int = 0):
    ar, ap, rec, cust = rows("ar_open.csv"), rows("ap_open.csv"), rows("recurring.csv"), rows("customers.csv")
    late = {c["customer"]: int(c["avg_days_late"]) for c in cust}
    receipts = defaultdict(Decimal)   # (week, category) -> amount
    payments = defaultdict(Decimal)
    rdetail, pdetail = [], []
    for r in ar:
        if lost_customer and r["customer"] == lost_customer:
            continue
        due = d(r["due_date"])
        expected = max(due + timedelta(days=late.get(r["customer"], 0) + slow_days), start)
        w = week_index(expected, start)
        if 0 <= w < weeks:
            receipts[(w, "Customer receipts")] += Decimal(r["amount_aed"])
            rdetail.append((r["invoice"], r["customer"], r["due_date"], expected.isoformat(), w + 1, Decimal(r["amount_aed"])))
    for b in ap:
        due = max(d(b["due_date"]) + timedelta(days=ap_stretch), start)
        w = week_index(due, start)
        if 0 <= w < weeks:
            payments[(w, b["category"])] += Decimal(b["amount_aed"])
            pdetail.append((b["bill"], b["supplier"], b["category"], due.isoformat(), w + 1, Decimal(b["amount_aed"])))
    for x in rec:
        amt = Decimal(x["amount_aed"])
        if x["frequency"] == "weekly":
            for w in range(weeks):
                (receipts if x["direction"] == "in" else payments)[(w, x["category"])] += amt
        elif x["frequency"] == "monthly":
            day = int(x["day_or_date"])
            m = start.replace(day=1)
            for _ in range(5):
                try:
                    dt = m.replace(day=day)
                except ValueError:
                    dt = (m.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
                w = week_index(dt, start)
                if 0 <= w < weeks and dt >= start:
                    (receipts if x["direction"] == "in" else payments)[(w, x["category"])] += amt
                m = (m.replace(day=28) + timedelta(days=4)).replace(day=1)
        elif x["frequency"] == "once":
            dt = d(x["day_or_date"])
            w = week_index(dt, start)
            if 0 <= w < weeks and dt >= start:
                (receipts if x["direction"] == "in" else payments)[(w, x["category"])] += amt
    rcats = sorted({k[1] for k in receipts})
    pcats = sorted({k[1] for k in payments})
    closing, cash = [], opening_cash
    for w in range(weeks):
        cash += sum(receipts[(w, c)] for c in rcats) - sum(payments[(w, c)] for c in pcats)
        closing.append(cash)
    return dict(receipts=receipts, payments=payments, rcats=rcats, pcats=pcats, closing=closing, rdetail=rdetail, pdetail=pdetail)


def head(ws, r, n):
    for c in range(1, n + 1):
        cell = ws.cell(row=r, column=c)
        cell.font = Font(bold=True, color=IVORY)
        cell.fill = HEAD
        cell.alignment = Alignment(horizontal="center" if c > 1 else "left", vertical="center")


def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x


def write(start: date, opening: Decimal, min_cash: Decimal, out: Path, client="Sample client"):
    weeks = 13
    base = build(start, opening)
    scen = {
        "Base": base,
        "Collections 14 days slower": build(start, opening, slow_days=14),
        "Largest customer lost": build(start, opening, lost_customer="Al Noor Trading LLC"),
        "Supplier terms stretched 7 days": build(start, opening, ap_stretch=7),
    }
    wb = Workbook()
    ws = wb.active
    ws.title = "Forecast"
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"13 week cash forecast: {client}"
    ws["A1"].font = Font(size=18, bold=True, color=INK, name="Georgia")
    low = min(base["closing"])
    lw = base["closing"].index(low) + 1
    ws["A2"] = (f"Lowest point AED {low:,.0f} in week {lw}" + (f", below the AED {min_cash:,.0f} minimum. " if low < min_cash else ". ")
                + f"Closing position AED {base['closing'][-1]:,.0f}. Built from open receivables, open payables and the recurring schedule.")
    ws["A2"].font = Font(size=11, italic=True, color=BRASS)
    ws["A3"] = f"Week 1 starts {start.isoformat()}. Receipts dated by each customer's average days late from history, not by the due date. Read only."
    ws["A3"].font = Font(size=9, italic=True, color=MUTED)
    r = 5
    ws.cell(row=r, column=1, value="AED")
    for w in range(weeks):
        ws.cell(row=r, column=2 + w, value=f"Wk {w + 1} ({(start + timedelta(days=7 * w)).strftime('%d %b')})")
        ws.cell(row=r, column=2 + w).alignment = Alignment(wrap_text=True, horizontal="center")
    ws.cell(row=r, column=2 + weeks, value="13 wk total")
    head(ws, r, weeks + 2)

    def line(label, values, bold=False, fill=None, fmt="#,##0;(#,##0);-"):
        nonlocal r
        r += 1
        ws.cell(row=r, column=1, value=label).font = Font(bold=bold, color=INK)
        for w, v in enumerate(values):
            c = ws.cell(row=r, column=2 + w, value=float(v))
            c.number_format = fmt
            c.font = Font(bold=bold, color=INK)
            if fill:
                c.fill = fill
        if values:
            c = ws.cell(row=r, column=2 + weeks, value=float(sum(values)) if label not in ("Opening cash", "Closing cash", "Headroom vs minimum") else float(values[-1]))
            c.number_format = fmt
            c.font = Font(bold=True, color=INK)
        for col in range(1, weeks + 3):
            ws.cell(row=r, column=col).border = BORDER
        if fill:
            ws.cell(row=r, column=1).fill = fill

    openings = [opening] + base["closing"][:-1]
    line("Opening cash", openings, bold=True, fill=BAND)
    r += 1
    ws.cell(row=r, column=1, value="Receipts").font = Font(bold=True, color=BRASS)
    tot_r = [Decimal(0)] * weeks
    for c in base["rcats"]:
        vals = [base["receipts"][(w, c)] for w in range(weeks)]
        tot_r = [a + b for a, b in zip(tot_r, vals)]
        line("   " + c, vals)
    line("Total receipts", tot_r, bold=True)
    r += 1
    ws.cell(row=r, column=1, value="Payments").font = Font(bold=True, color=BRASS)
    tot_p = [Decimal(0)] * weeks
    for c in base["pcats"]:
        vals = [base["payments"][(w, c)] for w in range(weeks)]
        tot_p = [a + b for a, b in zip(tot_p, vals)]
        line("   " + c, vals)
    line("Total payments", tot_p, bold=True)
    line("Net cash flow", [a - b for a, b in zip(tot_r, tot_p)], bold=True)
    line("Closing cash", base["closing"], bold=True, fill=BAND)
    close_row = r
    for w, v in enumerate(base["closing"]):
        if v < min_cash:
            ws.cell(row=close_row, column=2 + w).fill = RED
    line("Headroom vs minimum", [v - min_cash for v in base["closing"]])
    widths(ws, [30] + [12] * (weeks + 1))
    ws.freeze_panes = "B6"
    # chart
    ch = LineChart()
    ch.title = "Closing cash by week, base case"
    ch.height, ch.width = 7.5, 22
    ch.y_axis.title = "AED"
    ch.y_axis.number_format = "#,##0"
    data = Reference(ws, min_col=2, max_col=1 + weeks, min_row=close_row, max_row=close_row)
    ch.add_data(data, from_rows=True, titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=2, max_col=1 + weeks, min_row=5, max_row=5))
    ch.legend = None
    s0 = ch.series[0]
    s0.graphicalProperties.line.solidFill = BRASS
    s0.graphicalProperties.line.width = 28000
    s0.smooth = False
    ws.add_chart(ch, f"A{r + 3}")

    # ---- Scenarios
    sc = wb.create_sheet("Scenarios")
    sc.sheet_view.showGridLines = False
    sc["A1"] = "Closing cash by scenario"
    sc["A1"].font = Font(size=16, bold=True, color=INK, name="Georgia")
    sc["A2"] = "Same ledgers, one assumption changed at a time. The scenario that breaches the minimum first is the one to plan against."
    sc["A2"].font = Font(size=10, italic=True, color=MUTED)
    rr = 4
    sc.cell(row=rr, column=1, value="Scenario")
    for w in range(weeks):
        sc.cell(row=rr, column=2 + w, value=f"Wk {w + 1}")
    sc.cell(row=rr, column=2 + weeks, value="Low point")
    sc.cell(row=rr, column=3 + weeks, value="First breach")
    head(sc, rr, weeks + 3)
    for name, res in scen.items():
        rr += 1
        sc.cell(row=rr, column=1, value=name).font = Font(bold=True, color=INK)
        breach = next((w + 1 for w, v in enumerate(res["closing"]) if v < min_cash), None)
        for w, v in enumerate(res["closing"]):
            c = sc.cell(row=rr, column=2 + w, value=float(v))
            c.number_format = "#,##0;(#,##0);-"
            if v < min_cash:
                c.fill = RED
        sc.cell(row=rr, column=2 + weeks, value=float(min(res["closing"]))).number_format = "#,##0;(#,##0);-"
        sc.cell(row=rr, column=3 + weeks, value=f"Week {breach}" if breach else "None")
        for col in range(1, weeks + 4):
            sc.cell(row=rr, column=col).border = BORDER
    widths(sc, [34] + [11] * (weeks + 2))
    ch2 = LineChart()
    ch2.title = "Closing cash, four scenarios"
    ch2.height, ch2.width = 8, 24
    ch2.y_axis.number_format = "#,##0"
    data = Reference(sc, min_col=1, max_col=1 + weeks, min_row=5, max_row=4 + len(scen))
    ch2.add_data(data, from_rows=True, titles_from_data=True)
    ch2.set_categories(Reference(sc, min_col=2, max_col=1 + weeks, min_row=4, max_row=4))
    sc.add_chart(ch2, f"A{rr + 3}")

    # ---- Receipts detail
    rd = wb.create_sheet("Receipts detail")
    rd.sheet_view.showGridLines = False
    for c, h in enumerate(["Invoice", "Customer", "Due date", "Expected receipt", "Week", "AED"], 1):
        rd.cell(row=1, column=c, value=h)
    head(rd, 1, 6)
    for i, row in enumerate(sorted(base["rdetail"], key=lambda x: x[3]), 2):
        for c, v in enumerate(row, 1):
            cell = rd.cell(row=i, column=c, value=float(v) if c == 6 else v)
            cell.border = BORDER
            if c == 6:
                cell.number_format = "#,##0"
    widths(rd, [16, 30, 12, 16, 8, 14])
    # ---- Payments detail
    pd_ = wb.create_sheet("Payments detail")
    pd_.sheet_view.showGridLines = False
    for c, h in enumerate(["Bill", "Supplier", "Category", "Due date", "Week", "AED"], 1):
        pd_.cell(row=1, column=c, value=h)
    head(pd_, 1, 6)
    for i, row in enumerate(sorted(base["pdetail"], key=lambda x: x[3]), 2):
        for c, v in enumerate(row, 1):
            cell = pd_.cell(row=i, column=c, value=float(v) if c == 6 else v)
            cell.border = BORDER
            if c == 6:
                cell.number_format = "#,##0"
    widths(pd_, [16, 30, 22, 12, 8, 14])
    # ---- Assumptions
    asm = wb.create_sheet("Assumptions")
    asm.sheet_view.showGridLines = False
    lines = [("Week 1 start", start.isoformat()), ("Opening cash (AED)", float(opening)), ("Minimum cash (AED)", float(min_cash)),
             ("Receipt timing", "Due date plus the customer's average days late from the last 12 months"),
             ("Payment timing", "Bill due date; recurring items on their scheduled day"),
             ("Recurring items", "From recurring.csv: payroll, rent, VAT payment, corporate tax payment, subscriptions, retainer income"),
             ("Excluded", "Receipts already overdue by more than 90 days (treated as doubtful), receipts expected beyond week 13"),
             ("Scenario: collections slower", "Every customer pays 14 days later than history"),
             ("Scenario: largest customer lost", "All open invoices from Al Noor Trading LLC removed"),
             ("Scenario: supplier terms stretched", "Every bill paid 7 days after its due date")]
    for i, (k, v) in enumerate(lines, 1):
        asm.cell(row=i, column=1, value=k).font = Font(bold=True, color=INK)
        asm.cell(row=i, column=2, value=v)
        asm.cell(row=i, column=1).fill = BAND
    widths(asm, [34, 90])
    for w in wb.worksheets:
        w.page_setup.orientation = "landscape"
        w.page_setup.fitToWidth = 1
        w.page_setup.fitToHeight = 0
        w.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(out)
    return base, scen


if __name__ == "__main__":
    out = HERE / "out" / "cash-forecast-13w.xlsx"
    out.parent.mkdir(exist_ok=True)
    base, scen = write(date(2026, 9, 14), Decimal("185000"), Decimal("100000"), out)
    print("closing:", [f"{float(v):,.0f}" for v in base["closing"]])
    for k, v in scen.items():
        print(f"{k:34} low {float(min(v['closing'])):>12,.0f}  end {float(v['closing'][-1]):>12,.0f}")
    print("->", out)
