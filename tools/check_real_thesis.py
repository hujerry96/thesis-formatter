"""真實論文修正前後對比"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from core.analyzer import FormatAnalyzer
from core.fixer import FormatFixer
from collections import Counter

a = FormatAnalyzer('thesis_zh.yaml')
f = FormatFixer('thesis_zh.yaml')

print('=== BEFORE ===')
before = a.analyze('論文(最終版)-昱翔 - 複製.docx')
sev = Counter(i.get('severity', '?') for i in before)
types = Counter(i.get('type', '?') for i in before)
print(f'Total: {len(before)} ({dict(sev)})')
for t, c in types.most_common():
    print(f'  {t}: {c}')

print()
print('=== FIX ===')
fixes = f.fix_document('論文(最終版)-昱翔 - 複製.docx', 'test_output/real_thesis_fixed_v2.docx')
for fx in fixes:
    print(f'  - [{fx["type"]}] {fx["message"][:120]}')

print()
print('=== AFTER ===')
after = a.analyze('test_output/real_thesis_fixed_v2.docx')
sev2 = Counter(i.get('severity', '?') for i in after)
types2 = Counter(i.get('type', '?') for i in after)
print(f'Total: {len(after)} ({dict(sev2)})')
for t, c in types2.most_common():
    print(f'  {t}: {c}')
if len(after) < 30:
    for i in after:
        print(f'  - {i.get("message", "?")[:120]}')