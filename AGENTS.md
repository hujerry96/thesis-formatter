# 論文格式修正工具

## 專案結構

- `main.py` - 主程式入口
- `core/analyzer.py` - 格式分析器
- `core/fixer.py` - 格式修正器
- `core/rule_engine.py` - 規則引擎（YAML 配置）
- `gui/main_gui.py` - tkinter 圖形界面
- `rules/` - YAML 格式規則配置（thesis_zh.yaml）
- `requirements.txt` - Python 依賴
- `build_exe.py` - 打包成 exe 腳本

## 用途

自動修正 Word 論文格式錯誤，支援清華大學中英文論文格式。

## 開發注意

- 使用 Python + tkinter
- 規則引擎透過 YAML 配置定義格式規範
- 修改規則時編輯 `rules/` 下的 YAML 檔案
- 新增功能時同步更新 `requirements.txt`
