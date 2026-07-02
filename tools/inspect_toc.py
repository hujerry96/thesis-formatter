import sys; sys.path.insert(0, '.')
from docx import Document
import re

src = r'test_output\test_paper_B_complex_修正.docx'
doc = Document(src)

qn = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
heading_styles = {'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4',
                  'toc 1', 'toc 2', 'toc 3', 'toc 4'}

print("=== 檢查 heading/toc 段落文字樣本 ===")
for i, para in enumerate(doc.paragraphs):
    sname = para.style.name if para.style else ''
    if sname in heading_styles and para.text.strip():
        txt = para.text.strip()[:80]
        has_hyphen = '-' in txt
        mark = ' <-- 仍有連字號' if has_hyphen else ''
        print(f'  [{sname}] {txt}{mark}')
