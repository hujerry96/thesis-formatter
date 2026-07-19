# 論文格式修正工具 / Thesis Format Fixer

自動修正 Word 論文格式錯誤，支援碩士論文格式規範。
Automatically fix Word thesis formatting errors, supporting master's thesis format specs.

## 注意事項 / Notes

本工具根據規則對特定格式問題進行自動修正，若原始檔案排版過於混亂（例如樣式大量走樣、手動格式與樣式混用、多層巢狀異常等），可能難以達到完全理想的成效。建議先備份原始檔案，修正後再手動檢查。
This tool applies rule-based automatic fixes to specific formatting issues. If the source document is heavily disorganized (e.g. broken styles, mixed manual and styled formatting, deeply nested anomalies), results may be imperfect. Always back up the original file first and review the output manually.

## 功能 / Features

- **版面設定 / Layout**：頁邊距、裝訂邊、紙張大小 / Page margins, binding gutter, paper size
- **字型樣式 / Fonts**：標題樣式定義、字型統一（中文/英文） / Heading style definitions, unified fonts (Chinese/English)
- **行距段落 / Spacing**：行距 1.5 倍、段落間距、首行縮排 / 1.5 line spacing, paragraph spacing, first-line indent
- **頁碼設定 / Page Numbers**：前言羅馬數字、正文阿拉伯數字、起始頁碼 / Roman numerals for front matter, Arabic for body, start number
- **目錄 / TOC**：TOC 樣式修正、超連結格式清除 / TOC style fix, hyperlink format cleanup
- **標題 / Headings**：Heading 1-3 分頁設定、編號轉換（"第X章" ↔ "X."） / Heading 1-3 page breaks, numbering conversion
- **圖表標號 / Figures & Tables**：編號格式（圖 1-1 → 圖 1.1）、置中對齊 / Numbering format, centering
- **引用 / References**：交叉參照 REF 欄位、參考文獻編號列表 / Cross-reference REF fields, reference lists
- **封面 / Cover**：區塊間距調整、關鍵字可設定 / Cover spacing, configurable keywords
- **格式分析 / Analysis**：自動掃描文件問題並列出可修正項目 / Auto-scan and list fixable issues

## 安裝 / Install

```bash
pip install -r requirements.txt
```

需求：Python 3.8+，python-docx、lxml、PyYAML、tkinter（內建）
Requirements: Python 3.8+, python-docx, lxml, PyYAML, tkinter (built-in)

## 使用方式 / Usage

### GUI 模式 / GUI Mode

```bash
python main.py
```

拖曳 Word 檔案到視窗，或點擊「選取檔案」。
Drag a Word file into the window, or click "Select File".

介面語言可透過視窗右上角的下拉選單切換（中文 / English），選擇會自動記住。
The UI language can be switched via the dropdown at the top-right of the window (中文 / English); the choice is remembered.

### CLI 模式 / CLI Mode

```python
from core.analyzer import FormatAnalyzer
from core.fixer import FormatFixer

# 分析 / Analyze
analyzer = FormatAnalyzer('rules/thesis_zh.yaml')
issues = analyzer.analyze('論文.docx')
for i in issues:
    print(f"[{i['type']}] {i['message']}")

# 修正 / Fix
fixer = FormatFixer('rules/thesis_zh.yaml')
results = fixer.fix_document('論文.docx', '論文_修正.docx')
for r in results:
    print(f"[{r['type']}] {r['message']}")
```

## 專案結構 / Project Structure

```
├── main.py              # 主程式入口（tkinter GUI） / Entry point (tkinter GUI)
├── core/
│   ├── analyzer.py      # 格式分析器 / Format analyzer
│   ├── fixer.py        # 格式修正器 / Format fixer
│   └── rule_engine.py   # 規則引擎（讀取 YAML） / Rule engine (reads YAML)
├── gui/
│   ├── main_gui.py      # tkinter 圖形介面 / tkinter GUI
│   ├── template_dialog.py # 模板管理對話框 / Template manager dialog
│   └── i18n.py         # 多國語言模組 / Internationalization module
├── rules/
│   └── thesis_zh.yaml   # 中文格式規則 / Chinese format rules
├── build_exe.py         # PyInstaller 打包腳本 / PyInstaller build script
└── requirements.txt     # Python 依賴 / Python dependencies
```

## 規則配置 / Rule Configuration

編輯 `rules/` 下的 YAML 檔案可自訂：
Edit the YAML files under `rules/` to customize:

- 頁邊距、紙張大小、裝訂邊 / Page margins, paper size, binding gutter
- 中英文字型及大小 / Chinese & English fonts and sizes
- 行距倍數、段落間距 / Line spacing factor, paragraph spacing
- 首行縮排值 / First-line indent value
- 標題層級設定 / Heading level settings
- 圖表標號格式 / Figure & table numbering format
- 頁碼樣式（羅馬/阿拉伯） / Page number style (Roman/Arabic)
- 封面偵測關鍵字 / Cover detection keywords

## 打包為 EXE / Build EXE

```bash
python build_exe.py
```

產生 `dist/論文格式修正工具.exe`。
Produces `dist/Thesis Format Fixer.exe`.

## 開發 / Development

擴充步驟 / Extension steps:

1. 在 `core/analyzer.py` 新增 `_check_xxx()` 方法 / Add a `_check_xxx()` method in `core/analyzer.py`
2. 在 `core/fixer.py` 新增 `_fix_xxx()` 方法 / Add a `_fix_xxx()` method in `core/fixer.py`
3. 在 `__init__` 註冊到 `fix_methods` 列表 / Register it in the `fix_methods` list in `__init__`
4. 在 `rules/thesis_zh.yaml` 加入相關設定 / Add related settings in `rules/thesis_zh.yaml`
