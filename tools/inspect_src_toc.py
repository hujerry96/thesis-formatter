import sys; sys.path.insert(0, '.')
from docx import Document
import re

src = r'test_papers\test_paper_A_general.docx'
doc = Document(src)

print("=== 原始檔案中 toc 樣式段落 ===")
for i, para in enumerate(doc.paragraphs):
    sname = para.style.name if para.style else ''
    if 'toc' in sname.lower() or 'heading' in sname.lower():
        txt = para.text.strip()[:100]
        print(f'[{i:4d}] [{sname}] {repr(txt[:60])}')
