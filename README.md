# NPL Deal Review App

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## What changed in this version

- Keeps the title compact so screen space goes to the deal review itself.
- Places the Asset Manager directly beside the deal context in the review header and queue labels.
- Shows a denser deal presentation with more datapoints visible at once.
- Adds a dedicated Presentation mode tab for cleaner Teams screen-sharing.
- Splits comments into:
  - Previous Weekly Comment (read-only, from the uploaded workbook)
  - This Week Comment (editable in the app)
  - Export Preview for the final single Excel `Comments` cell
- Prepends the current weekly comment back into the original `Comments` column on download so the exported workbook still looks like the original workbook layout.
- Adds standardized picklists for the most repetitive binary / trinary / status-based AM update fields.
- Keeps a bulk-edit grid for faster review sessions across many deals.

## Workflow

1. Upload the latest workbook.
2. Filter the review queue as needed.
3. Review individual deals in the Deal review or Presentation mode tab.
4. Enter this week’s updates and other AM manual fields.
5. Download the updated workbook.

## Notes

- The app preserves the original workbook structure and writes updates back into the existing deal sheet.
- Excel is set to recalculate when the downloaded workbook is opened.
- The app defaults to the workbook-visible deal rows but can include workbook-hidden rows when needed.
