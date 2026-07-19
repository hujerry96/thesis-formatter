[English](README.md) | [繁體中文](README.zh.md)

# Thesis Format Fixer

Automatically fix Word thesis formatting errors, supporting master's thesis format specs.

## Notes

This tool applies rule-based automatic fixes to specific formatting issues. If the source document is heavily disorganized (e.g. broken styles, mixed manual and styled formatting, deeply nested anomalies), results may be imperfect. Always back up the original file first and review the output manually.

## Features

- **Layout**: Page margins, binding gutter, paper size
- **Fonts**: Heading style definitions, unified fonts (Chinese/English)
- **Spacing**: 1.5 line spacing, paragraph spacing, first-line indent
- **Page Numbers**: Roman numerals for front matter, Arabic for body, start number
- **TOC**: TOC style fix, hyperlink format cleanup
- **Headings**: Heading 1-3 page breaks, numbering conversion ("Chapter X" <-> "X.")
- **Figures & Tables**: Numbering format (Figure 1-1 -> Figure 1.1), centering
- **References**: Cross-reference REF fields, reference lists
- **Cover**: Cover spacing, configurable keywords
- **Analysis**: Auto-scan and list fixable issues

## Install

```bash
pip install -r requirements.txt
```

Requirements: Python 3.8+, python-docx, lxml, PyYAML, tkinter (built-in)

## Usage

### GUI Mode

```bash
python main.py
```

Drag a Word file into the window, or click "Select File".

The UI language can be switched via the dropdown at the top-right of the window (中文 / English); the choice is remembered.

### CLI Mode

```python
from core.analyzer import FormatAnalyzer
from core.fixer import FormatFixer

# Analyze
analyzer = FormatAnalyzer('rules/thesis_zh.yaml')
issues = analyzer.analyze('thesis.docx')
for i in issues:
    print(f"[{i['type']}] {i['message']}")

# Fix
fixer = FormatFixer('rules/thesis_zh.yaml')
results = fixer.fix_document('thesis.docx', 'thesis_fixed.docx')
for r in results:
    print(f"[{r['type']}] {r['message']}")
```

## Project Structure

```
├── main.py              # Entry point (tkinter GUI)
├── core/
│   ├── analyzer.py      # Format analyzer
│   ├── fixer.py        # Format fixer
│   └── rule_engine.py   # Rule engine (reads YAML)
├── gui/
│   ├── main_gui.py      # tkinter GUI
│   ├── template_dialog.py # Template manager dialog
│   └── i18n.py         # Internationalization module
├── rules/
│   └── thesis_zh.yaml   # Chinese format rules
├── build_exe.py         # PyInstaller build script
└── requirements.txt     # Python dependencies
```

## Rule Configuration

Edit the YAML files under `rules/` to customize:

- Page margins, paper size, binding gutter
- Chinese & English fonts and sizes
- Line spacing factor, paragraph spacing
- First-line indent value
- Heading level settings
- Figure & table numbering format
- Page number style (Roman/Arabic)
- Cover detection keywords

## Build EXE

```bash
python build_exe.py
```

Produces `dist/Thesis Format Fixer.exe`.

## Development

Extension steps:

1. Add a `_check_xxx()` method in `core/analyzer.py`
2. Add a `_fix_xxx()` method in `core/fixer.py`
3. Register it in the `fix_methods` list in `__init__`
4. Add related settings in `rules/thesis_zh.yaml`
