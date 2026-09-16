# HR ERP

A Streamlit application for monthly attendance processing, payroll calculation, and individual PDF salary-slip generation.

## Features

- Upload one Excel workbook containing attendance and salary data.
- Normalise `HH.MM`, standard Excel time, and text time values.
- Apply documented late-mark, half-day, early-leave, overtime, and salary rules.
- Review a daily operations dashboard, monthly attendance trends, payroll costs, employee risk signals, and a prioritised HR action queue.
- Review record-level validation exceptions before finalising payroll.
- Download a payroll workbook and individual PDF salary slips.

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy to Streamlit Community Cloud

Choose this repository's `main` branch and set **Main file path** to `streamlit_app.py`.

## Rules implemented

- Daily pay is monthly salary divided by 30.
- The first four late marks are allowed; every later late mark deducts 0.5 day's pay.
- A check-in after 11:00 or duration of four hours or less is a half day.
- The first two early leaves are allowed and excess instances are flagged without a financial deduction.
- Overtime is awarded in completed 3.5-hour blocks after 18:00: each block earns 0.5 day's pay.

See the in-app Rules applied panel for the complete set of operational assumptions.
