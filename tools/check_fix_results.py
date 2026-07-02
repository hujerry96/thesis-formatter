"""檢查修正前後的問題數對比"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from core.analyzer import FormatAnalyzer
from core.fixer import FormatFixer
from collections import Counter

analyzer = FormatAnalyzer('thesis_zh.yaml')
fixer = FormatFixer('thesis_zh.yaml')

papers = [
    ('test_papers/test_paper_A_general.docx', '論文 A'),
    ('test_papers/test_paper_B_complex.docx', '論文 B'),
    ('test_papers/test_paper_C_edge.docx', '論文 C'),
]

for path, label in papers:
    print(f'=== {label} ===')
    
    # Fix
    out = f'test_output/{Path(path).stem}_fixed_v2.docx'
    fixes = fixer.fix_document(path, out)
    print(f'  Fixes applied: {len(fixes)}')
    for f in fixes:
        print(f'    - [{f["type"]}] {f["message"][:100]}')
    
    # Re-analyze the fixed output
    remaining = analyzer.analyze(out)
    sev = Counter(i.get('severity', '?') for i in remaining)
    types = Counter(i.get('type', '?') for i in remaining)
    print(f'  After fix - remaining issues: {len(remaining)} ({dict(sev)})')
    for t, c in types.most_common():
        print(f'    {t}: {c}')
    print()