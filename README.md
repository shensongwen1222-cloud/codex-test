# Local Excel Report Merger

A lightweight static web app for merging multiple Excel reports into a single summary workbook. It uses pure HTML, CSS, JavaScript, and SheetJS in the browser—no backend or upload server required.

## Features

- Upload multiple `.xlsx`, `.xls`, or `.csv` files
- Merge all rows from every worksheet in each workbook into one table
- Add `Source File` and `Source Sheet` columns for traceability
- Generate per-column summary statistics:
  - Non-empty values
  - Blank values
  - Unique values
  - Numeric sum, average, minimum, and maximum
- Preview the first 100 merged rows
- Export `summary.xlsx` with `Merged Data` and `Summary Stats` sheets

## Run locally

Open `index.html` directly in a modern browser, or serve the folder with any static file server:

```bash
python3 -m http.server 8000
```

Then visit <http://localhost:8000>.

> Note: Excel parsing and export are powered by SheetJS loaded in the browser from the official SheetJS CDN. File contents stay in the browser and are not sent to a backend.
