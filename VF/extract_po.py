"""
Extract PO header + line-item + subline (size/color/qty) data
from the VF Corporation / Vans purchase order PDF.

HOW TO RUN:
    python3 extract_po.py

Requires:
    pip install pdfplumber
"""
import pdfplumber
import re
import json
import glob
import sys

# --- Auto-detect the PDF in the current folder ---
pdf_files = glob.glob("*.pdf")
if not pdf_files:
    print("ERROR: No PDF file found in this folder.")
    print("Please copy your PO PDF into the same folder as this script, then run again.")
    sys.exit(1)
elif len(pdf_files) > 1:
    print("Multiple PDFs found in this folder:")
    for i, f in enumerate(pdf_files, 1):
        print(f"  {i}. {f}")
    choice = input("Enter the number of the PDF to use: ")
    PDF_PATH = pdf_files[int(choice) - 1]
else:
    PDF_PATH = pdf_files[0]

print(f"Using PDF: {PDF_PATH}\n")


with pdfplumber.open(PDF_PATH) as pdf:
    full_text = "\n".join(page.extract_text() for page in pdf.pages)

# ---------- Line items (top-level, e.g. 600090369200002) ----------
line_pattern = re.compile(
    r"(\d{15})\s+([A-Z0-9]+)\s+([A-Z0-9 \-/]+?)\s+([\d,]+)\s+(?:PIECE|EACH)\s+([\d.]+)\s+([\d,]+\.\d{2})"
)
lines = []
for m in line_pattern.finditer(full_text):
    lines.append({
        "po_line_no": m.group(1),
        "style": m.group(2),
        "description": m.group(3).strip(),
        "order_qty": int(m.group(4).replace(",", "")),
        "unit_price": float(m.group(5)),
        "line_cost": float(m.group(6).replace(",", "")),
    })

# ---------- Sublines (size / color / qty / UPC per size) ----------
subline_pattern = re.compile(
    r"(\d{12,13})\s+([A-Z0-9]+)\s+[A-Z0-9 \-/]+?\s+([\d,]+)\s+([\d.]+)\s*\n"
    r"[A-Z0-9 %\-/]+?\s*\n"
    r"UnitOfMeasureCode\s+(?:PC|EA)\s+Color\s+([A-Z /]+?)\s*\n"
    r"Size\s+(\S+(?:\s\S+)?)\s+Dimension\s+1\s*\n"
    r"SKU Number\s+(\S+)\s+UPC Number\s+(\d+)\s*\n"
    r"Packing Method\s+CASE\s+Items Per Outer\s+(\d+)\s*\nPack"
)
sublines = []
for m in subline_pattern.finditer(full_text):
    sublines.append({
        "subline_no": m.group(1),
        "style": m.group(2),
        "qty": int(m.group(3).replace(",", "")),
        "unit_cost": float(m.group(4)),
        "color": m.group(5).strip(),
        "size": m.group(6).strip(),
        "sku": m.group(7),
        "upc": m.group(8),
        "items_per_outer_pack": int(m.group(9)),
    })

# ---------- CRD (Customer/Brand Requested delivery date) - one per line ----------
crd_matches = re.findall(r"Brand Requested CRD\s+(\d{4}-\d{2}-\d{2})", full_text)
for i, l in enumerate(lines):
    l["crd"] = crd_matches[i] if i < len(crd_matches) else None

# ---------- Destination country (from the Ship To address block) ----------
# NOTE: best-effort pattern based on this PDF's layout - the "Ship To" box is a
# rotated/vertical text column, so pdfplumber reads it out of order. If this
# ever returns None (or the wrong value) on a differently-formatted PO,
# open raw_pdf_text.txt (from inspect_pdf_text.py) and search for the
# country name manually, then adjust this regex.
country_match = re.search(
    r"([A-Z]{2,}(?:\s[A-Z]{2,})*)\s*\n[A-Za-z]?\s*a\s*\na\s*\nBrand Buyer Code",
    full_text
)
destination_country = country_match.group(1) if country_match else None
for l in lines:
    l["destination_country"] = destination_country

print("=== PO LINES FOUND ===")
for l in lines:
    print(l)

print(f"\n=== SUBLINES FOUND: {len(sublines)} ===")
for s in sublines:
    print(s)

# Saved in the SAME folder you run this script from
with open("po_extracted.json", "w") as f:
    json.dump({"lines": lines, "sublines": sublines}, f, indent=2)

print("\nSaved -> po_extracted.json")
