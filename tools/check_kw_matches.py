import sys; sys.path.insert(0, '.')
from docx import Document
import re

KW_RE = re.compile(r'^(關鍵詞|關鍵字|Key words|Keywords)\s*[:：]?\s*', re.IGNORECASE)

src = r'論文(最終版)-昱翔 - 複製.docx'
doc = Document(src)
for i, para in enumerate(doc.paragraphs):
    txt = para.text.strip()
    if KW_RE.match(txt):
        print(f'[{i:4d}] {repr(txt[:80])}')
