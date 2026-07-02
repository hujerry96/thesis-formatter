#!/usr/bin/env python3
"""
論文格式修正工具
自動修正 Word 論文格式錯誤
"""

import sys
from pathlib import Path

def main():
    print("=" * 50)
    print("論文格式修正工具 v2.0")
    print("=" * 50)
    
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        template_key = sys.argv[2] if len(sys.argv) > 2 else 'zh'
        
        if Path(input_file).exists():
            from core.rule_engine import TEMPLATES
            from core.analyzer import FormatAnalyzer
            from core.fixer import FormatFixer
            
            input_path = Path(input_file)
            output_path = input_path.parent / f"{input_path.stem}_格式修正{input_path.suffix}"
            
            rule_file = TEMPLATES.get(template_key, 'thesis_zh.yaml')
            
            print(f"分析檔案: {input_path}")
            print(f"模板: {template_key}")
            analyzer = FormatAnalyzer(rule_file)
            issues = analyzer.analyze(str(input_path))
            
            print(f"\n發現 {len(issues)} 個格式問題:")
            for issue in issues:
                print(f"  [{issue.get('severity', 'info')}] {issue.get('message', '')}")
            
            print(f"\n修正檔案並儲存至: {output_path}")
            fixer = FormatFixer(rule_file)
            fix_results = fixer.fix_document(str(input_path), str(output_path))
            for fix in fix_results:
                print(f"  [{fix.get('type', '')}] {fix.get('message', '')}")
            print("完成！")
        else:
            print(f"錯誤: 檔案不存在 {input_file}")
    else:
        print("啟動圖形界面...")
        from gui.main_gui import ThesisFormatterGUI
        app = ThesisFormatterGUI()
        app.mainloop()

if __name__ == "__main__":
    main()
