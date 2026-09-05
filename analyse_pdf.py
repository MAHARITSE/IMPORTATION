import pdfplumber, os

PDF_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "PDF")

files = [
    "14-08-2026 AFG ASSURANCES 1050265 SALFATU LOT94_082026-ocr.pdf",
    "[2025-07-24] 07-07-2025 SARO 1080364 SALFATU LOT118_072025-ocr.pdf",
    "[2026-01-05] 17-12-2025 BFV-SG 1128267 SALFATU LOT85_122025-ocr.pdf",
]
for fname in files:
    path = os.path.join(PDF_DIR, fname)
    print(f"\n{'='*60}")
    print(f"FILE: {fname}")
    print('='*60)
    with pdfplumber.open(path) as pdf:
        print(f"Pages: {len(pdf.pages)}")
        # page 2 ou 1
        page = pdf.pages[1] if len(pdf.pages) > 1 else pdf.pages[0]
        print(page.extract_text()[:3000])
