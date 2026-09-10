"""Sample ledgers for a UAE services business with ~AED 5m revenue: open AR, open AP, recurring schedule, customer behaviour."""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(7)
OUT = Path(__file__).resolve().parent / "sample"
OUT.mkdir(exist_ok=True)
START = date(2026, 9, 14)

customers = [("Al Noor Trading LLC", 18), ("Sharjah Textiles FZE", 6), ("Gulf Marine Services", 31), ("Ajman Foods Co", 12),
             ("RAK Ceramics Distribution", 3), ("Fujairah Logistics", 24), ("Blue Coast Cafe", 0), ("Nordic Imports AB", 9),
             ("Desert Rose Clinics", 15), ("Emirates Steel Fabrication", 40)]
with open(OUT / "customers.csv", "w", newline="") as f:
    w = csv.writer(f); w.writerow(["customer", "avg_days_late"]); w.writerows(customers)

ar = []
n = 1200
for k in range(48):
    cust, _ = random.choice(customers)
    issue = START - timedelta(days=random.randint(1, 70))
    due = issue + timedelta(days=random.choice([30, 30, 45, 60]))
    amt = random.choice([6500, 9800, 12500, 18000, 22500, 28000, 36000, 44000]) + random.randint(0, 900)
    if cust == "Al Noor Trading LLC":
        amt = int(amt * 1.8)
    ar.append([f"INV-{n + k}", cust, issue.isoformat(), due.isoformat(), amt])
with open(OUT / "ar_open.csv", "w", newline="") as f:
    w = csv.writer(f); w.writerow(["invoice", "customer", "issue_date", "due_date", "amount_aed"]); w.writerows(ar)

suppliers = [("Etisalat", "Utilities and telecoms"), ("DEWA", "Utilities and telecoms"), ("Zoho Corp", "Software"), ("Microsoft", "Software"),
             ("Al Futtaim Motors", "Vehicles"), ("Emirates Post", "Operations"), ("Gulf Office Supplies", "Operations"), ("Freelance developers", "Contractors"),
             ("Dubai Insurance", "Insurance"), ("Flydubai", "Travel")]
ap = []
for k in range(40):
    sup, cat = random.choice(suppliers)
    due = START + timedelta(days=random.randint(-5, 85))
    amt = random.choice([1200, 2400, 3800, 6500, 9000, 14500, 22000]) + random.randint(0, 400)
    ap.append([f"BILL-{700 + k}", sup, due.isoformat(), amt, cat])
with open(OUT / "ap_open.csv", "w", newline="") as f:
    w = csv.writer(f); w.writerow(["bill", "supplier", "due_date", "amount_aed", "category"]); w.writerows(ap)

recurring = [
    ["Payroll", "Payroll", 168000, "monthly", "25", "out"],
    ["Office rent (quarterly cheque)", "Rent", 96000, "once", "2026-10-01", "out"],
    ["VAT payment Q3", "Tax", 74000, "once", "2026-10-28", "out"],
    ["Corporate tax payment", "Tax", 61000, "once", "2026-09-30", "out"],
    ["Bank charges and card fees", "Bank", 1800, "weekly", "", "out"],
    ["Fuel and transport", "Operations", 2600, "weekly", "", "out"],
    ["Retainer: Desert Rose Clinics", "Retainer income", 18500, "monthly", "5", "in"],
    ["Retainer: Blue Coast Cafe", "Retainer income", 6500, "monthly", "5", "in"],
    ["Health insurance renewal", "Insurance", 42000, "once", "2026-11-15", "out"],
]
with open(OUT / "recurring.csv", "w", newline="") as f:
    w = csv.writer(f); w.writerow(["name", "category", "amount_aed", "frequency", "day_or_date", "direction"]); w.writerows(recurring)
print("ar", len(ar), "ap", len(ap), "recurring", len(recurring))
