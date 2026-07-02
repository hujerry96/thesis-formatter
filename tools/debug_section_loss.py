"""逐步執行每個 fix 步驟，找出哪步丟失 sectPr"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from docx.oxml.ns import qn
from pathlib import Path

INPUT = r"論文(最終版)-昱翔 - 複製.docx"
OUTDIR = r"tools\debug_sections"
Path(OUTDIR).mkdir(exist_ok=True)

def count_sections(doc):
    body = doc.element.body
    sectPrs = body.findall(qn('w:sectPr'))
    psectPrs = body.findall(f'.//{qn("w:pPr")}/{qn("w:sectPr")}')
    return len(sectPrs), len(psectPrs)

def count_sections_via_api(doc):
    return len(doc.sections)

doc = Document(INPUT)
s1, p1 = count_sections(doc)
api1 = count_sections_via_api(doc)
print(f"原始文件: body sectPr={s1}, pPr/sectPr={p1}, API sections={api1}")

from core.fixer import FormatFixer
fixer = FormatFixer()

fix_methods = [
    ('page_margins', fixer._fix_page_margins),
    ('heading_style_defs', fixer._fix_heading_style_defs),
    ('direct_font_formatting', fixer._fix_direct_font_formatting),
    ('heading_indent', fixer._remove_heading_indent),
    ('heading_page_break', fixer._fix_heading_page_break),
    ('special_sections_page_break', fixer._fix_special_sections_page_break),
    ('line_spacing', fixer._fix_line_spacing),
    ('first_line_indent', fixer._fix_first_line_indent),
    ('numbering_xml', fixer._fix_numbering_xml),
    ('empty_paragraphs', fixer._fix_empty_paragraphs),
    ('toc_hyperlinks', fixer._fix_toc_hyperlinks),
    ('caption_numbering', fixer._fix_caption_numbering),
    ('caption_alignment', fixer._fix_caption_alignment),
    ('toc_field_codes', fixer._fix_toc_field_codes),
    ('references_numbered_list', fixer._convert_references_to_numbered_list),
    ('cross_references', fixer._fix_cross_references),
    ('inline_citations_xref', fixer._convert_inline_citations_to_xrefs),
    ('toc_style', fixer._fix_toc_style),
    ('cover_tabs', fixer._fix_cover_tabs),
    ('page_numbers', fixer._fix_page_numbers),
]

for i, (name, method) in enumerate(fix_methods):
    try:
        method(doc)
    except Exception as e:
        print(f"  {name}: ERROR - {e}")
    
    s, p = count_sections(doc)
    api = count_sections_via_api(doc)
    status = "*** SECTION LOST ***" if api < api1 else "OK"
    print(f"Step {i+1:2d} {name:40s} body={s} pPr={p} API={api} {status}")
    
    outpath = os.path.join(OUTDIR, f"step_{i+1:02d}_{name}.docx")
    doc.save(outpath)
    
    if api < api1:
        api1 = api
        print(f"  >>> Section lost after {name}! Saved to {outpath}")
