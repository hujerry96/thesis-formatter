from core.analyzer import FormatAnalyzer
a = FormatAnalyzer()
results = a.analyze(r'test_output/real_thesis_fixed_v4.docx')
for r in results:
    msg = r['message']
    print(f"{r.get('type','?')}: {msg[:120]}")
