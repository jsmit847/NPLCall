# NPL Deal Review Streamlit App (v3)

This version is tuned around the completed weekly workbook that was uploaded.

## What changed

- Compact screen-share friendly layout with a much smaller title/header footprint
- Asset Manager displayed directly beside the deal name in every detail / presentation view
- Dynamic support for both workbook variants observed so far:
  - `Loan List 3.27`
  - `Loan List 4.3`
- Better handling of the weekly comment workflow:
  - prior comment history stays visible
  - AM enters only **This Week Comment**
  - export prepends the dated current-week note back into the workbook `Comments` cell
- Controlled picklists based on the workbook values actually being used this week
- Queue helpers for fast review:
  - next / previous
  - next open deal
  - next blank current-week comment
  - queue filters for missing comments, discussion items, and flagged rows
- Presentation mode optimized for Teams sharing
- Bulk update grid with read-only prior comment and editable current-week comment
- Data-quality flags to surface rows missing key fields or using inconsistent status values

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Upload flow

1. Upload the latest workbook
2. Review the queue and filter to the AM or deal set you want
3. Enter this week's updates in **Review** or **Bulk update**
4. Use **Presentation mode** during Teams screen share
5. Download the workbook export when done

## Export behavior

The workbook export patches the active deal sheet in place so the original workbook structure, formatting, and sheet layout are preserved.

## Notes

- The app uses the workbook's currently active deal layout and property detail sheet automatically.
- The `Comments` export is built from:
  - current-week entry date
  - this-week comment typed in the app
  - existing historical comments already in the workbook
