# Thirteen week cash forecast

Builds a 13 week cash forecast from open receivables, open payables, the recurring schedule and customer payment behaviour, and runs it under four scenarios. Built and tested on sample data by [Finamatik](https://finamatik.com/work/cash-flow-forecast-13-weeks).

## What it does

Every receipt is dated by how each customer actually pays (average days late from history), every payment by its due date or its recurring rule (payroll, rent, VAT and corporate tax, subscriptions, retainers). Opening and closing cash per week, weeks below the minimum flagged, and a detail sheet behind every week so any number can be traced to a line.

Scenarios from the same ledgers: base, collections 14 days slower, largest customer lost, supplier terms stretched 7 days, each with its first breach week and the headroom against the minimum.

## Run

```
python3 make_sample.py      # sample/ar_open.csv, ap_open.csv, recurring.csv, customers.csv
python3 forecast.py         # out/cash-forecast-13w.xlsx
```

Python 3.11, openpyxl. Inputs are four small CSVs that any accounting system can export; the assumptions sheet in the workbook says exactly what was and was not included.

MIT licence, copyright Finamatik Business Solutions FZE LLC. Questions and production use: info@finamatik.com.
