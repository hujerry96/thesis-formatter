import sys; sys.path.insert(0, '.')
from docx import Document
from lxml import etree

qn = lambda x: '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' + x

src = r'論文(最終版)-昱翔 - 複製.docx'
doc = Document(src)

# 找 SDT 中的目錄段落
for child in doc.element.body:
    if not child.tag.endswith('}sdt'):
        continue
    sdt = child.find(qn('w:sdtContent'))
    if sdt is None:
        continue
    sdt_text = ''.join(t.text or '' for t in sdt.findall('.//' + qn('t')))
    if '第一章' not in sdt_text and '1-1' not in sdt_text:
        continue
    print("=== 找到目錄 SDT ===")
    # 找每個 run 的內容
    for r in sdt.findall('.//' + qn('r')):
        t = r.find(qn('t'))
        if t is not None and t.text:
            txt = t.text.strip()
            if txt and len(txt) < 80:
                print(f'  run: {repr(t.text)}')
    break
