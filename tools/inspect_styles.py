import sys; sys.path.insert(0, '.')
from docx import Document
from collections import Counter

src = r'論文(最終版)-昱翔 - 複製.docx'
doc = Document(src)

style_counts = Counter()
for i, para in enumerate(doc.paragraphs):
    sname = para.style.name if para.style else 'None'
    style_counts[sname] += 1

print("=== 段落樣式分布 ===")
for s, c in style_counts.most_common(50):
    print(f'  {s:30s} x{c}')
