# Deal KPI Review & AM Update App

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## What it does

- Upload the latest workbook first.
- Detects the main deal sheet and the related property detail sheet.
- Shows each deal with KPI metrics, core datapoints, and the assigned Asset Manager.
- Lets you edit the AM manual update fields in either a deal-by-deal form or a bulk edit grid.
- Downloads an updated workbook while preserving the original workbook layout and formatting.

## Notes

- The app keeps the workbook structure and Excel filters intact.
- If any formula-backed manual cells are edited, the downloaded workbook is set to recalculate when opened in Excel.
