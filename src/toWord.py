import sys
import re
from docx import Document
from docx.oxml.ns import qn

def inputWord(text, file):
    doc = Document()
    p = doc.add_paragraph()
    run = p.add_run(text)
    
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    
    doc.save(f"{file}.docx")   


def changeWord(file_path):
    doc = Document(file_path)
    
    # 定義數學函數關鍵字
    math_functions = ['sin', 'cos', 'tan', 'log']

    for paragraph in doc.paragraphs:

        old_text = paragraph.text.replace('-', '−')
        
        paragraph.text = ""
        
        parts = re.split(r'([a-zA-Z]+)', old_text)
        
        for part in parts:
            if not part:
                continue
                
            run = paragraph.add_run(part)
            
            run.font.name = 'Times New Roman'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            
            if part.islower() and part.isalpha():
                if part in math_functions:
                    run.italic = False  
                else:
                    run.italic = True   
            else:   
                run.italic = False

    try:
            doc.save(file_path)
    except PermissionError:
        print("\n" + "-" * 30)
        print(f"請關閉檔案{file_path}後，再重新執行一次程式。")
        print("-" * 30 + "\n")
        sys.exit(1)
    