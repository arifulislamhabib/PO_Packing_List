"""
Diagnostic script — shows the RAW text pdfplumber extracts from the PDF,
BEFORE any regex parsing happens. Use this to check what data is
actually being read, and to debug regex patterns if extraction ever
misses something.

HOW TO RUN:
    python3 inspect_pdf_text.py

It will:
  1. Auto-detect the PDF in this folder (same as extract_po.py)
  2. Print the raw text page-by-page to the terminal
  3. Save the full raw text to raw_pdf_text.txt so you can open it
     in a text editor and search through it freely
"""
import pdfplumber
import glob
import sys

pdf_files = glob.glob("*.pdf")
if not pdf_files:
    print("ERROR: No PDF file found in this folder.")
    sys.exit(1)
elif len(pdf_files) > 1:
    print("Multiple PDFs found:")
    for i, f in enumerate(pdf_files, 1):
        print(f"  {i}. {f}")
    choice = input("Enter the number of the PDF to use: ")
    PDF_PATH = pdf_files[int(choice) - 1]
else:
    PDF_PATH = pdf_files[0]

print(f"Reading: {PDF_PATH}\n")

with pdfplumber.open(PDF_PATH) as pdf:
    print(f"Total pages: {len(pdf.pages)}\n")
    all_text = []
    for i, page in enumerate(pdf.pages, start=1):
        text = page.extract_text()
        all_text.append(text or "")
        print(f"{'='*70}")
        print(f" PAGE {i}")
        print(f"{'='*70}")
        print(text)
        print()

# Save it all to a plain text file you can open/search in any editor
with open("raw_pdf_text.txt", "w") as f:
    f.write("\n".join(all_text))

print(f"\nFull raw text also saved to: raw_pdf_text.txt")
print("Open that file in any text editor to search through it, or")
print("compare it line-by-line against what extract_po.py's regex captures.")
