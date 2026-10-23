# Synthetic portfolio reporting packs — Q2 2026 (test set v1)

All companies and numbers are fictional. Built to test a quarterly-review pipeline:
messy packs in → standard KPIs out, scored against an answer key.

## Contents

```
fx_rates_to_GBP.csv
set_A_infrastructure/      5 infrastructure portfolio companies (ISQ-style)
    packs/A1..A5/          one folder per company = one quarterly pack
    answer_key.csv         ground truth: extracted inputs + derived KPIs
    discrepancies.csv      every planted problem + expected pipeline behaviour
set_B_asset_managers/      5 asset-management firms (AUM-based KPIs)
    packs/B1..B5/  answer_key.csv  discrepancies.csv
```

Every pack contains a metadata file (`metadata.json` or `metadata.yaml`). Metadata keys,
date formats and units deliberately differ between packs.

| Pack | Company | Files | Main problems |
|---|---|---|---|
| A1 | Halden Flex Power (GBP) | xlsx, pdf | merged headers, hidden row, covenant breach, narrative contradiction |
| A2 | Corvane Fibre (EUR) | csv, docx, yaml | kEUR, decimal comma, synonym label, missing capex plan |
| A3 | Frostline Cold (EUR) | 2× pdf (French) | French labels, capex ×10 outlier, Q1 restatement |
| A4 | Meridian Bus (GBP) | xlsx, email txt | fiscal year Apr–Mar, £000, Excel serial dates, missing headcount, LTIR spike |
| A5 | Verdan Recycling (GBP) | json, docx | H1 YTD only (derive Q2), US dates, tight covenant |
| B1 | Ashcombe Capital (GBP) | xlsx, pdf | own-funds threshold breach, narrative contradiction, % as decimals |
| B2 | Lindqvist & Moor (EUR) | csv, txt, yaml | bn and m mixed in one file, missing 5y performance |
| B3 | Harbourline (USD) | pdf | fiscal year ends Sept, US dates, AUM rollforward doesn't tie |
| B4 | Kestrel Point (GBP) | docx (narrative) | redemptions ×10 in table, "positive flows" claim vs negative flows |
| B5 | Solenne Gestion (EUR) | 2× pdf (French) | French labels, opening AUM restated, space thousands separator |

## Conventions the answer key uses

- **Target period:** calendar Q2 2026 (1 Apr – 30 Jun 2026). Fiscal labels must be mapped.
- **Values:** millions, in the company's own currency (`expected_value_native`).
- **GBP column:** flows (revenue, EBITDA, capex, cash flow, fund flows, fees, costs) at the
  Q2 average rate; balances (net debt, closing AUM) at the 30 Jun closing rate; opening AUM at
  the 31 Mar closing rate. Rates in `fx_rates_to_GBP.csv`.
- **Ratios** are computed in native currency, so FX never affects them.
- **Missing values** are blank in the key and should come out as null + a flag, never 0.
- **Restated figures:** the key uses the restated number.
- **Outliers:** the key holds the value that reconciles; the pipeline should flag rather than silently fix.

## KPI formulas used in the answer key

Set A
- EBITDA margin % = EBITDA / revenue × 100
- vs budget % = actual / budget − 1; QoQ % = Q2 / Q1 − 1
- Cash conversion % = operating cash flow / EBITDA × 100
- Leverage (x) = net debt / LTM EBITDA; covenant headroom = covenant max − leverage (≤0.5x = tight, <0 = breach)

Set B
- AUM growth % = closing / opening − 1
- Organic growth % = net flows / opening AUM (quarterly; annualised = ×4)
- Market contribution % = market movement / opening AUM
- Fee rate (bps, annualised) = net revenue × 4 / average AUM × 10,000
- Operating margin % = (revenue − costs) / revenue; cost-income % = costs / revenue

These definitions are the starting point for your KPI dictionary.

## How to score a pipeline run

1. Output the same long format as `answer_key.csv` (company_id, metric, scenario, value).
2. Join on company_id + metric + scenario; a value is correct within ±0.5% (or exact for nulls).
3. Separately check that every row in `discrepancies.csv` produced a flag.

Regenerate or vary the data with `generate_packs.py` (change the numbers in the INFRA / AM dicts).
