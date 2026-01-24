import fitz  # PyMuPDF
import shutil
import os

def pdf_to_picture(pdf_path, start_page=1, end_page=0):         #可選擇頁數
    pdf = fitz.open(pdf_path)
    end_page = len(pdf) if end_page == 0 else end_page
    
    
    if os.path.exists("picture"):
        shutil.rmtree("picture")  
    os.makedirs("picture")
    
    
    for page_num in range(start_page - 1, end_page):
        page = pdf.load_page(page_num)
        pix = page.get_pixmap(dpi=300)
        pix.save(f"picture/page_{page_num + 1}.png")
    # print(len(pdf))
    pdf.close()