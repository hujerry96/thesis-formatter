import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn
import xml.etree.ElementTree as ET

doc = Document(r'test_output/real_thesis_fixed_v5.docx')

for i, sec in enumerate(doc.sections):
    print(f'\n=== Section {i} footer ===')
    xml = ET.tostring(sec.footer._element, encoding='unicode')
    print(xml)
