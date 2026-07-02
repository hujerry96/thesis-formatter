import sys; sys.path.insert(0, '.')
from docx import Document
import re, os
from core.fixer import FormatFixer

src = r'論文(最終版)-昱翔 - 複製.docx'
out = r'test_output\test_heading_number.docx'

fixer = FormatFixer()
fixer.input_path = src
fixer.output_path = out

doc = Document(src)
result = fixer._fix_heading_numbering(doc)
print('_fix_heading_numbering result:', result)

tgt_re = re.compile(r'(\d)-(\d)')
heading_styles = {'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4',
                  'toc 1', 'toc 2', 'toc 3', 'toc 4'}
count = 0
qn = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
for para in doc.paragraphs:
    sname = para.style.name if para.style else ''
    if sname not in heading_styles:
        continue
    if not para.text.strip():
        continue
    for r in para._element.findall('.//' + qn + 'r'):
        t = r.find(qn + 't')
        if t is not None and t.text and '-' in t.text:
            new_text = tgt_re.sub(r'\1.\2', t.text)
            if new_text != t.text:
                count += 1
                print(f'  [{sname}] "{t.text}" -> "{new_text}"')

print(f'轉換後仍含連字號的 heading/toc 段落: {count}')
