[English](README.md) | [繁體中文](README.zh.md)

# 論文格式修正工具

自動修正 Word 論文格式錯誤，支援碩士論文格式規範。

## 注意事項

本工具根據規則對特定格式問題進行自動修正，若原始檔案排版過於混亂（例如樣式大量走樣、手動格式與樣式混用、多層巢狀異常等），可能難以達到完全理想的成效。建議先備份原始檔案，修正後再手動檢查。

## 功能

- **版面設定**：頁邊距、裝訂邊、紙張大小
- **字型樣式**：標題樣式定義、字型統一（中文/英文）
- **行距段落**：行距 1.5 倍、段落間距、首行縮排
- **頁碼設定**：前言羅馬數字、正文阿拉伯數字、起始頁碼
- **目錄**：TOC 樣式修正、超連結格式清除
- **標題**：Heading 1-3 分頁設定、編號轉換（"第X章" ↔ "X."）
- **圖表標號**：編號格式（圖 1-1 → 圖 1.1）、置中對齊
- **引用**：交叉參照 REF 欄位、參考文獻編號列表
- **封面**：區塊間距調整、關鍵字可設定
- **格式分析**：自動掃描文件問題並列出可修正項目

## 安裝

```bash
pip install -r requirements.txt
```

需求：Python 3.8+，python-docx、lxml、PyYAML、tkinter（內建）

## 使用方式

### GUI 模式

```bash
python main.py
```

拖曳 Word 檔案到視窗，或點擊「選取檔案」。

介面語言可透過視窗右上角的下拉選單切換（中文 / English），選擇會自動記住。

### CLI 模式

```python
from core.analyzer import FormatAnalyzer
from core.fixer import FormatFixer

# 分析
analyzer = FormatAnalyzer('rules/thesis_zh.yaml')
issues = analyzer.analyze('論文.docx')
for i in issues:
    print(f"[{i['type']}] {i['message']}")

# 修正
fixer = FormatFixer('rules/thesis_zh.yaml')
results = fixer.fix_document('論文.docx', '論文_修正.docx')
for r in results:
    print(f"[{r['type']}] {r['message']}")
```

## 專案結構

```
├── main.py              # 主程式入口（tkinter GUI）
├── core/
│   ├── analyzer.py      # 格式分析器
│   ├── fixer.py        # 格式修正器
│   └── rule_engine.py   # 規則引擎（讀取 YAML）
├── gui/
│   ├── main_gui.py      # tkinter 圖形介面
│   ├── template_dialog.py # 模板管理對話框
│   └── i18n.py         # 多國語言模組
├── rules/
│   └── thesis_zh.yaml   # 中文格式規則
├── build_exe.py         # PyInstaller 打包腳本
└── requirements.txt     # Python 依賴
```

## 規則配置

編輯 `rules/` 下的 YAML 檔案可自訂：

- 頁邊距、紙張大小、裝訂邊
- 中英文字型及大小
- 行距倍數、段落間距
- 首行縮排值
- 標題層級設定
- 圖表標號格式
- 頁碼樣式（羅馬/阿拉伯）
- 封面偵測關鍵字

## 打包為 EXE

```bash
python build_exe.py
```

產生 `dist/論文格式修正工具.exe`。

## 開發

擴充步驟：

1. 在 `core/analyzer.py` 新增 `_check_xxx()` 方法
2. 在 `core/fixer.py` 新增 `_fix_xxx()` 方法
3. 在 `__init__` 註冊到 `fix_methods` 列表
4. 在 `rules/thesis_zh.yaml` 加入相關設定
