# Beyond the Report: JSE mining risk dashboard (APFA802)

Interactive Streamlit dashboard for Valterra Platinum, Impala Platinum, Sibanye-Stillwater,
Gold Fields and Exxaro Resources, FY2021 to FY2025.

## Run it

```
pip install -r requirements.txt
streamlit run app.py
```

It opens at http://localhost:8501. To share a link for the panel, push this folder to a GitHub
repository and deploy it free on Streamlit Community Cloud (share.streamlit.io), with app.py as the entry point.

## Folder contents

- `app.py` - the dashboard (10 pages).
- `esgrq.py` - ESG Reporting Quality index (paper section 9 / Appendix A): 17 items scored 0-4, six
  dimensions, comparability, assurance and validation (second-coder agreement, weighted kappa).
  Run `python esgrq.py` to print the scores.
- `data/esgrq_coding.xlsx` - the coding workbook the group fills in (sheets: Read me, Coding,
  Assurance, Comparability). Every scored row needs a source document and page.
- The source report PDFs live one folder up (`Nonto/`), not in this app folder, to keep deploys small.
- `loader.py` - reads the five workbooks in `data/`, standardises line items, computes ratios,
  the risk scorecard and the data-quality checks. Run `python loader.py` to print everything.
- `data/*.xlsx` - the group's financial-statement workbooks (renamed, unchanged).
- `data/esg_data.csv` - create this from the template on the ESG page; it loads automatically.
- `data/feedback.csv` - created when testers submit the form on "Test the dashboard".

## Conventions

- Amounts in million. Gold Fields reports in US$; the sidebar converts it to rand with
  editable average USD/ZAR rates. Ratios do not depend on currency.
- Costs, capex, dividends, finance costs and tax expense are stored as positive numbers.
- EBIT = profit before tax + finance costs, so all five firms are measured the same way.
- Impala's year ends 30 June; the others 31 December.

## Before submission

1. Fill in `report_page` for every extracted line (Data and quality page, download the CSV).
2. Fix the High-severity items listed on the Data and quality page.
3. `data/esgrq_coding.xlsx` holds an AI-drafted first coding of the 2025 reports (coder column says so).
   A group member must verify every score against the cited page. Sibanye-Stillwater is not yet coded.
   Second coder: fill `data/esgrq_recheck_blind.xlsx` (no first-coder scores shown) without opening the
   coding workbook, then run `python esgrq.py merge`. Regenerate the blind sheet with `python esgrq.py blind`.
   Declare the AI-assisted coding in the AI-use declaration.
4. Capture ESG indicators into the template, with source document and page for each value.
5. Verify the USD/ZAR rates.
