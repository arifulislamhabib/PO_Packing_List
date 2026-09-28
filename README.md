# PO → Packing List Converter

A Streamlit web app that reads a VF/Vans-style Purchase Order PDF and generates
a packing-list-style Excel workbook — one sheet per PO line — with live formulas
for cartons, weights, totals, excess/short, and CBM.

Everything runs in the browser. No terminal steps for the end user, no manual
data entry, no external services.

---

## Table of Contents

- [Features](#features)
- [Screenshots](#screenshots)
- [Requirements](#requirements)
- [Quick Start (Local)](#quick-start-local)
- [Docker](#docker)
- [How It Works](#how-it-works)
- [Project Structure](#project-structure)
- [Usage](#usage)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

---

## Features

### PDF Extraction

- Parses VF/Vans-style PO PDFs using `pdfplumber`
- Extracts **PO lines**: PO number, style, description, order qty, unit price, line cost
- Extracts **sublines**: size, color, SKU, UPC, quantity, Items Per Outer Pack
- Reads CRD (Cancel/Request Date) and destination country from the header
- Silently skips malformed blocks so one bad subline doesn't kill the whole parse

### Interactive Preview (Streamlit)

- **Editable carton grid** — change size quantities or TOTAL CTN directly in cells
- **Add / remove carton rows** with `＋` and `❌` buttons (perfectly aligned grid)
- **Packing qty MIN / MAX** override — user's MAX replaces the PDF's pack size
- Live recalculation of CTN PCS, TOTAL PCS, gross/net weight, and totals
- **Reset** button to restore rows to PDF defaults
- **Excess / Short banner** — green if over-ship, red if short
- Per-size, per-color, per-UPC layout with wider Color and UPC columns
- Order-quantity-by-size summary table (horizontal scroll on narrow screens)

### Excel Export (`.xlsx`)

- Fully **center-aligned** document-style layout
- Title box, size-spec table (SIZE / N.W. / N.N.W. / EMPTY CTN), info block,
  main carton table, TOTAL row, summary block, footer, signature lines
- **Merged COLOR cells** for consecutive same-color runs
- **MIXED** rows highlighted yellow
- Live Excel formulas (SUMPRODUCT, IF, IFERROR) so the workbook stays editable
  in Excel/LibreOffice
- Frozen header rows + hidden gridlines for a clean print view
- Two Grand Total columns (gross / net) merged vertically in the summary

### PDF Export (`.pdf`)

- Generated with `reportlab` (landscape A4)
- Info table + main carton table + summary table + footer + signature
- Same numbers as the xlsx, ready for printing

### Quality-of-life

- **Timestamped filenames** — downloads don't overwrite previous revisions
- **Download extracted JSON** — audit or re-run the pipeline offline
- **Auto-clear session state** when a new PDF is uploaded
- Clean, modern UI theme (Inter font, subtle shadows, card-based metrics)

---

## Screenshots

> Add your own screenshots to a `docs/` folder and link them here.
