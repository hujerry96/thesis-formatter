import sys; sys.path.insert(0, '.')
from docx import Document
from lxml import etree

qn = lambda x: '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' + x

src = r'論文(最終版)-昱翔 - 複製.docx'
doc = Document(src)

# 找標題為「目錄」的段落索引
toc_idx = None
for i, para in enumerate(doc.paragraphs):
    if para.text.strip() == '目錄':
        toc_idx = i
        break

# 檢查第一個 sdt 內是否有目錄
body = doc.element.body
sdts = body.findall('.//' + qn('sdt'))
print(f'SDT 總數: {len(sdts)}')

# 找 sdt 內含 TOC 或目錄文字的
for idx, sdt in enumerate(sdts[:30]):
    sdtc = sdt.find(qn('sdtContent'))
    if sdtc is None:
        continue
    text = ''.join(t.text or '' for t in sdtc.findall('.//' + qn('t')))
    if text.strip():
        print(f'\nSDT #{idx}: text={repr(text[:120])}')
        # 找段落樣式
        pstyles = sdtc.findall('.//' + qn('pStyle'))
        for ps in pstyles:
            print(f'  pStyle val={ps.get(qn("val"))}')
