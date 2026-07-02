from docx import Document
from docx.oxml.ns import qn

doc = Document(r'test_output/real_thesis_fixed_v4.docx')
for i, sec in enumerate(doc.sections):
    print(f'Section {i}: type={sec.start_type}')
    sp = sec._sectPr
    pgnum = sp.find(qn('w:pgNumType'))
    if pgnum is not None:
        print(f'  pgNumType start={pgnum.get(qn("w:start"))}, fmt={pgnum.get(qn("w:fmt"))}')
    footer = sec.footer
    print(f'  footer linked_to_previous={footer.is_linked_to_previous}')
    for j, p in enumerate(footer.paragraphs):
        runs = p._element.findall(qn('w:r'))
        flds = [r for r in runs if r.find(qn('w:fldChar')) is not None or r.find(qn('w:instrText')) is not None]
        print(f'  footer para {j}: {len(runs)} runs, {len(flds)} fld-related, text="{p.text}"')
