from docx import Document
from docx.oxml.ns import qn

print('=== ORIGINAL ===')
doc = Document('論文(最終版)-昱翔 - 複製.docx')
body = doc.element.body
children = list(body)
print(f'Section count: {len(doc.sections)}')
count = 0
for i, child in enumerate(children):
    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
    if tag == 'sectPr':
        print(f'  {i}: sectPr (section {count})')
        count += 1

print('\n=== FIXED ===')
doc2 = Document('test_output/real_thesis_fixed_v2.docx')
body2 = doc2.element.body
children2 = list(body2)
print(f'Section count: {len(doc2.sections)}')
count = 0
for i, child in enumerate(children2):
    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
    if tag == 'sectPr':
        print(f'  {i}: sectPr (section {count})')
        count += 1