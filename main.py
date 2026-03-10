import os
import cv2
import numpy as np
import fitz  # pip install pymupdf
import pandas as pd
from tqdm import tqdm  # pip install tqdm
from paddleocr import PPStructure, save_structure_res

# 1. Initialize for CPU
structure_engine = PPStructure(
    use_gpu=False,
    show_log=False,  # Set to False so the log doesn't mess up the progress bar
    layout=True,
    table=True,
    version='PP-OCRv4'
)

pdf_path = "final_pdf.pdf"
save_folder = "./output_results"
master_excel_path = os.path.join(save_folder, "Master_Tables_Extraction.xlsx")
os.makedirs(save_folder, exist_ok=True)

# 2. Open PDF
doc = fitz.open(pdf_path)
total_pages = len(doc)
all_tables_count = 0

print(f"🚀 Starting OCR on {total_pages} pages...")

# 3. Process with Progress Bar
with pd.ExcelWriter(master_excel_path, engine='openpyxl') as writer:
    # tqdm wraps the range to create the visual bar
    for i in tqdm(range(total_pages), desc="Processing PDF Pages", unit="page"):
        page = doc.load_page(i)

        # Convert page to image
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, 3)
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

        # Process the page
        result = structure_engine(img)
        save_structure_res(result, save_folder, f"page_{i}")

        # Extract Tables
        table_index = 0
        for region in result:
            if region['type'] == 'table':
                try:
                    table_html = region['res']['html']
                    df = pd.read_html(table_html)[0]
                    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)

                    sheet_name = f"P{i}_Table_{table_index}"
                    df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

                    all_tables_count += 1
                    table_index += 1
                except Exception:
                    continue

doc.close()

print(f"\n✅ Done! Found {all_tables_count} tables across {total_pages} pages.")
print(f"📂 Results: {save_folder}")
print(f"📊 Master Excel: {master_excel_path}")