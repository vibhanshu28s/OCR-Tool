import os
import cv2
import numpy as np
import fitz  # pip install pymupdf
import pandas as pd
from tqdm import tqdm
from paddleocr import PPStructure, save_structure_res

# 1. Initialize Engine
structure_engine = PPStructure(
    use_gpu=False,
    show_log=False,
    layout=True,
    table=True,
    version='PP-OCRv4'
)

pdf_path = "final_pdf.pdf"
save_folder = "./output_results"
master_excel_path = os.path.join(save_folder, "Master_Tables.xlsx")
master_text_path = os.path.join(save_folder, "Extracted_Content.md")

os.makedirs(save_folder, exist_ok=True)

doc = fitz.open(pdf_path)
all_text_content = []

print(f"🚀 Extracting Text and Tables from {len(doc)} pages...")

with pd.ExcelWriter(master_excel_path, engine='openpyxl') as writer:
    for i in tqdm(range(len(doc)), desc="Processing"):
        page = doc.load_page(i)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, 3)
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

        # Process Page
        result = structure_engine(img)

        page_text = f"## Page {i + 1}\n\n"

        table_count = 0
        for region in result:
            # TYPE A: Extract Tables to Excel
            if region['type'] == 'table':
                try:
                    df = pd.read_html(region['res']['html'])[0]
                    sheet_name = f"P{i + 1}_Tab_{table_count}"
                    df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

                    page_text += f"*[Table extracted to Excel sheet: {sheet_name}]*\n\n"
                    table_count += 1
                except:
                    continue

            # TYPE B: Extract Text (Title, Figure Caption, or Paragraph)
            else:
                # PaddleOCR stores text lines in 'res' for non-table regions
                region_text = ""
                for line in region['res']:
                    region_text += line['text'] + " "

                if region['type'] == 'title':
                    page_text += f"### {region_text.strip()}\n\n"
                else:
                    page_text += f"{region_text.strip()}\n\n"

        all_text_content.append(page_text)

# 5. Save all text to a Markdown file
with open(master_text_path, "w", encoding="utf-8") as f:
    f.write("# Document Extraction Report\n\n")
    f.writelines(all_text_content)

doc.close()
print(f"\n✅ Done! \n📊 Tables: {master_excel_path}\n📝 Text: {master_text_path}")