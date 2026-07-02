"""分析測試論文 A 的 issue 分布"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from core.analyzer import FormatAnalyzer
from collections import Counter

analyzer = FormatAnalyzer('thesis_zh.yaml')
issues = analyzer.analyze('test_papers/test_paper_A_general.docx')

sev = Counter(i.get('severity', 'unknown') for i in issues)
print('Severity breakdown:', dict(sev))

types = Counter(i.get('type', 'unknown') for i in issues)
print('\nType breakdown:')
for t, c in types.most_common(20):
    print(f'  {t}: {c}')

print('\nSample issues (first 15):')
for i in issues[:15]:
    print(f'  [{i.get("severity", "?")}] {i.get("type", "?")}: {i.get("message", "?")[:100]}')