import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn
import xml.etree.ElementTree as ET

doc = Document(r'test_output/real_thesis_fixed_v5.docx')

print(f'Total sections: {len(doc.sections)}')
for i, sec in enumerate(doc.sections):
    print(f'\n--- Section {i} ---')
    sp = sec._sectPr
    pgnum = sp.find(qn('w:pgNumType'))
    if pgnum is not None:
        print(f'  pgNumType: start={pgnum.get(qn("w:start"))}')
    footer = sec.footer
    print(f'  footer linked_to_previous={footer.is_linked_to_previous}')
    ftr = footer._element
    children = list(ftr)
    print(f'  footer children count: {len(children)}')
    sdt_count = len(ftr.findall(qn('w:sdt')))
    print(f'  SDT count: {sdt_count}')
    paras = list(ftr.iter('{' + qn('w:p').split('}')[0] + '}p'))
    print(f'  paragraph elements: {len(paras)}')
    for j, p in enumerate(paras):
        runs = p.findall(qn('w:r'))
        instrs = [r.find(qn('w:instrText')) for r in runs]
        instr_texts = []
        for it in instrs:
            if it is not None and it.text:
                instr_texts.append(it.text)
        print(f'  para {j}: {len(runs)} runs, instrTexts={instr_texts}')
