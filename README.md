# ATM & Card Usage Analysis: RBI Data (September 2026)

Who controls India's ATM and card infrastructure, and are cards used for **spending** or for **cash withdrawal**?
This project cleans the RBI's raw bank-wise report with Python, builds 34 KPIs and measures market concentration.

## Key findings

| Finding | Number |
|---|---|
| Total ATMs / CRMs | 209,761 (34.8% off-site) |
| SBI share of ATMs | 31.1% (65,209 ATMs) |
| Private banks: share of ATMs vs credit card spend | 33.7% vs **73.8%** |
| Public sector banks: share of ATMs vs credit card spend | 64.3% vs **20.4%** |
| HDFC share of credit card spend | 31.9% (top 5 banks: 79.0%) |
| Debit card value taken out as ATM cash | **86.5%** |
| Online share of credit card spend | 61.8% |
| Spend per card per month | about Rs 16,300 (credit) vs about Rs 330 (debit) |
| UPI QR code concentration (HHI) | 5,673 (highly concentrated) |

**In one line:** India has two card economies. Debit cards, held mostly through public sector banks, are mainly a cash tool. Credit cards, held mostly through private banks, are mainly a spending tool.

![Category share](output/charts/02_category_share.png)
![Cash dependence](output/charts/06_cash_dependence.png)
![Pareto](output/charts/10_pareto_cc_spend.png)

## Data

- **Source:** Reserve Bank of India, monthly *ATM, Acceptance Infrastructure and Card Statistics* (September 2026, provisional).
- **Size:** 63 banks x 26 metrics, in 5 categories (public, private, foreign, payment, small finance banks).
- **Level:** one row per bank, one month. Values are in Rs thousand (converted to crore in the analysis).

## Method

1. Read the Excel report with no header (the real header spans 7 merged rows).
2. Detect section-heading rows and bank rows with rules, and attach each bank to its category.
3. Convert to numeric and fill blanks with 0 (blank = bank reported none).
4. **Validate:** the ATM total from my cleaned data matches the sheet's own Total row (209,761).
5. Engineer 34 KPIs: spend per card, average ticket size, cash dependence, cards per ATM, e-com share, market shares.
6. Category summary built from sums (weighted ratios), not averages of bank ratios.
7. Concentration: top-N share and HHI.
8. Export a clean CSV, an Excel summary, 10 charts and an insights report.

## Project structure

```
rbi-atm-card-analysis/
├── README.md
├── requirements.txt
├── atm_card_analysis.py        <- the full analysis script
├── data/                       <- RBI source file
└── output/
    ├── charts/                 <- 10 PNG charts
    ├── key_insights.txt
    ├── analysis_summary.xlsx
    └── clean_bank_data.csv
```

## How to run

```bash
pip install -r requirements.txt
python atm_card_analysis.py
```

The script finds the Excel file in `data/` automatically. To use another file: `python atm_card_analysis.py "path/to/file.xlsx"`.

## Limitations

- One month of provisional, bank-reported data, so there is no trend.
- Cards only. UPI transactions are not included, so low debit card spend does not mean low digital payments.
- Some ratios use small denominators, so thresholds are applied (for example, at least 1 lakh credit cards for per-card rankings).
- Blank cells are treated as zero.

## Next steps

Add 12 or more months to study trends, add UPI data, build a Power BI or Streamlit dashboard, and cluster banks into segments.

## Tools

Python, pandas, NumPy, Matplotlib, openpyxl
