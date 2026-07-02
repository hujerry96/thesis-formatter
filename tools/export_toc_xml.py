import sys; sys.path.insert(0, '.')
from docx import Document
from lxml import etree

qn = lambda x: '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' + x

src = r'論文(最終版)-昱翔 - 複製.docx'
doc = Document(src)

with open(r'test_output\toc_xml.txt', 'w', encoding='utf-8') as f:
    for child in doc.element.body:
        if not child.tag.endswith('}sdt'):
            continue
        sdt = child.find(qn('w:sdtContent'))
        if sdt is None:
            continue
        sdt_text = ''.join(t.text or '' for t in sdt.findall('.//' + qn('t')))
        if '第一章' in sdt_text or '1-1' in sdt_text:
            f.write(etree.tostring(child, pretty_print=True, encoding='unicode'))
            print(f'SDT text has pattern, wrote XML ({len(sdt_text)} chars)')
            break

# 現在也看修正後的 SDT
with open(r'test_output\toc_xml_fixed.txt', 'w', encoding='utf-8') as f:
    from core.fixer import FormatFixer
    fixer = FormatFixer()
    doc2 = Document(src)
    fixer._fix_heading_numbering(doc2)
    for child in doc2.element.body:
        if not child.tag.endswith('}sdt'):
            continue
        sdt = child.find(qn('w:sdtContent'))
        if sdt is None:
            continue
        sdt_text = ''.join(t.text or '' for t in sdt.findall('.//' + qn('t')))
        if '第一章' in sdt_text:
            f.write(etree.tostring(child, pretty_print=True, encoding='unicode'))
            print(f'Fixed SDT text: {repr(sdt_text[:200])}')
            break
