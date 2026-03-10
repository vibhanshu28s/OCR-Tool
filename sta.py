import os
import cv2
import numpy as np
import fitz  # pip install pymupdf
from paddleocr import PPStructure, save_structure_res

# 1. Initialize for CPU
structure_engine = PPStructure(
    use_gpu=False,
    show_log=True,
    version='PP-OCRv4'
)

pdf_path = "final_pdf.pdf"
save_folder = "./output_results"
os.makedirs(save_folder, exist_ok=True)

# 2. Open PDF and iterate through pages manually
doc = fitz.open(pdf_path)

for i, page in enumerate(doc):
    # Convert page to image (pixmap) at 2x scale for better OCR accuracy
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))

    # Convert pixmap to numpy array
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, 3)

    # Convert RGB to BGR for OpenCV compatibility
    img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

    # 3. Process the individual page image
    # This prevents the 'list' object has no attribute 'shape' error
    result = structure_engine(img)

    # 4. Save results for the current page
    save_structure_res(result, save_folder, f"final_pdf_page_{i}")

doc.close()
print(f"Success! Results saved to {save_folder}")