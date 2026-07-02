import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn
import xml.etree.ElementTree as ET

doc = Document(r'論文(最終版)-昱翔 - 複製.docx')

# Dump footer1, footer2, footer3 raw xml
for i, sec in enumerate(doc.sections):
    print(f'\n=== Section {i} ===')
    ftr = sec.footer._element
    xml = ET.tostring(ftr, encoding='unicode')
    print(xml[:1500])
