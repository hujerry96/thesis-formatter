"""介面多國語言模組 / UI internationalization module.

ponytail: dict 查表取代硬編碼字串，避免引入 gettext 等外部依賴。
語言選擇寫入 config.ini，下次啟動記住。
"""

from pathlib import Path

CONFIG_PATH = Path(__file__).resolve().parent.parent / "config.ini"

DEFAULT_LANG = "zh"
LANGS = ("zh", "en")
LANG_NAMES = {"zh": "中文", "en": "English"}

_STRINGS = {
    "zh": {
        "app_title": "論文格式修正工具",
        "app_subtitle": "自動檢查並修正 Word 論文格式問題",
        "file_card_title": "選擇論文檔案",
        "file_card_desc": "載入需要修正格式的 Word 文件",
        "drop_hint": "將 .docx 檔案拖曳至此處，或點擊下方按鈕選擇",
        "drop_active": "放開以選取此檔案",
        "browse": "瀏覽…",
        "template_card_title": "格式模板",
        "template_card_desc": "選擇適用的論文格式規範",
        "manage_templates": "管理模板…",
        "clear_results": "清除結果",
        "run": "開始格式修正",
        "col_severity": "等級",
        "col_type": "類型",
        "col_message": "說明",
        "sev_error": "錯誤",
        "sev_warning": "警告",
        "sev_info": "提示",
        "sev_fixed": "修正",
        "builtin": "內建",
        "custom": "自訂",
        "status_analyzing": "正在分析與修正中…",
        "status_done": "完成！已儲存至 {}",
        "err_no_file": "錯誤",
        "err_no_file_msg": "請選擇有效的 .docx 檔案",
        "err_process": "處理錯誤",
        "err_process_msg": "模板「{}」處理失敗:\n{}",
        "file_dialog_title": "選擇論文檔案",
        "filetype_docx": "Word 文件",
        "filetype_all": "所有檔案",
        "tm_title": "模板管理",
        "tm_heading_name": "模板名稱",
        "tm_heading_key": "識別碼",
        "tm_heading_source": "來源",
        "tm_add": "新增",
        "tm_edit": "編輯",
        "tm_delete": "刪除",
        "tm_import": "匯入",
        "tm_export": "匯出",
        "tm_duplicate": "複製",
        "tm_close": "關閉",
        "tm_status_created": "已建立自訂模板: {}",
        "tm_status_saved": "已儲存模板: {}",
        "tm_status_deleted": "已刪除模板: {}",
        "tm_status_imported": "已匯入模板: {}",
        "tm_status_exported": "已匯出至: {}",
        "tm_status_duplicated": "已複製模板: {}",
        "tm_err_empty_key": "請輸入識別碼",
        "tm_err_empty_name": "請輸入模板名稱",
        "tm_err_title": "錯誤",
        "info": "提示",
        "tm_hint_select": "請先選擇一個模板",
        "tm_hint_builtin_noedit": "內建模板無法編輯，請複製為自訂模板後再編輯",
        "tm_hint_builtin_nodelete": "內建模板無法刪除",
        "tm_confirm_delete": "確認刪除",
        "tm_confirm_delete_msg": "確定刪除自訂模板「{}」？",
        "tm_import_title": "選擇要匯入的模板",
        "tm_filetype_template": "模板檔案",
        "tm_filetype_yaml": "YAML 檔案",
        "tm_filetype_json": "JSON 檔案",
        "tm_export_title": "匯出模板",
        "tm_err_load": "無法載入模板: {}",
        "tm_exists_title": "模板已存在",
        "tm_exists_overwrite": "自訂模板已存在，是否覆蓋？",
        "tm_overwrite_ok": "已覆蓋匯入模板",
        "tm_add_title": "新增自訂模板",
        "tm_dup_title": "複製模板",
        "tm_based_on": "以「{}」為基礎",
        "tm_label_key": "識別碼",
        "tm_label_name": "模板名稱",
        "tm_cancel": "取消",
        "tm_create": "建立",
        "tm_duplicate_btn": "複製",
        "te_title_new": "新增模板",
        "te_title_edit": "編輯模板",
        "te_err_load": "無法載入編輯器:\n{}",
        "te_label_name": "模板名稱",
        "te_label_desc": "描述",
        "te_tab_paper": " 紙張與頁面 ",
        "te_tab_spacing": " 行距與段落 ",
        "te_tab_headings": " 標題樣式 ",
        "te_tab_misc": " 內文·頁碼·圖表 ",
        "te_paper": "紙張",
        "te_size": "尺寸",
        "te_orient": "方向",
        "te_margin_body": "內文邊距 (cm)",
        "te_margin_cover": "封面邊距 (cm)",
        "te_bind": "裝訂",
        "te_bind_extra": "額外左邊距 (cm)",
        "te_top": "上",
        "te_bottom": "下",
        "te_left": "左",
        "te_right": "右",
        "te_ls": "行距",
        "te_ls_type": "類型",
        "te_ls_value": "數值",
        "te_ps": "段落間距 (pt)",
        "te_ps_before": "段前",
        "te_ps_after": "段後",
        "te_indent": "首行縮排",
        "te_indent_enabled": "啟用",
        "te_indent_cm": "數值 (cm)",
        "te_indent_chars": "字元數",
        "te_h_title": "摘要/致謝/目錄",
        "te_h1": "標題 1",
        "te_h2": "標題 2",
        "te_h3": "標題 3",
        "te_h4": "標題 4",
        "te_zh_font": "中文字體",
        "te_en_font": "英文字體",
        "te_font_size": "大小",
        "te_bold": "粗體",
        "te_yes": "是",
        "te_no": "否",
        "te_align": "對齊",
        "te_before": "段前 (pt)",
        "te_after": "段後 (pt)",
        "te_page_break": "段前分頁",
        "te_body": "內文樣式",
        "te_fm": "前置頁頁碼",
        "te_fm_style": "格式",
        "te_fm_pos": "位置",
        "te_fm_align": "對齊",
        "te_bm": "正文頁碼",
        "te_bm_start": "起始頁碼",
        "te_fig": "圖片格式",
        "te_fig_num": "編號格式",
        "te_fig_caption": "標題在下方",
        "te_tab": "表格格式",
        "te_tab_caption": "標題在上方",
        "te_save": "儲存",
        "te_err_name": "驗證錯誤",
        "te_err_name_msg": "模板名稱不能為空白",
    },
    "en": {
        "app_title": "Thesis Format Fixer",
        "app_subtitle": "Automatically check and fix Word thesis formatting",
        "file_card_title": "Select Thesis File",
        "file_card_desc": "Load the Word document to be reformatted",
        "drop_hint": "Drag a .docx file here, or click the button below",
        "drop_active": "Release to select this file",
        "browse": "Browse…",
        "template_card_title": "Format Template",
        "template_card_desc": "Choose the applicable thesis format spec",
        "manage_templates": "Manage Templates…",
        "clear_results": "Clear Results",
        "run": "Start Formatting",
        "col_severity": "Level",
        "col_type": "Type",
        "col_message": "Description",
        "sev_error": "Error",
        "sev_warning": "Warning",
        "sev_info": "Info",
        "sev_fixed": "Fixed",
        "builtin": "Built-in",
        "custom": "Custom",
        "status_analyzing": "Analyzing and fixing…",
        "status_done": "Done! Saved to {}",
        "err_no_file": "Error",
        "err_no_file_msg": "Please select a valid .docx file",
        "err_process": "Processing Error",
        "err_process_msg": 'Template "{}" failed:\n{}',
        "file_dialog_title": "Select Thesis File",
        "filetype_docx": "Word Document",
        "filetype_all": "All Files",
        "tm_title": "Template Manager",
        "tm_heading_name": "Name",
        "tm_heading_key": "Key",
        "tm_heading_source": "Source",
        "tm_add": "Add",
        "tm_edit": "Edit",
        "tm_delete": "Delete",
        "tm_import": "Import",
        "tm_export": "Export",
        "tm_duplicate": "Duplicate",
        "tm_close": "Close",
        "tm_status_created": "Custom template created: {}",
        "tm_status_saved": "Template saved: {}",
        "tm_status_deleted": "Template deleted: {}",
        "tm_status_imported": "Template imported: {}",
        "tm_status_exported": "Exported to: {}",
        "tm_status_duplicated": "Template duplicated: {}",
        "tm_err_empty_key": "Please enter a key",
        "tm_err_empty_name": "Please enter a template name",
        "tm_err_title": "Error",
        "info": "Info",
        "tm_hint_select": "Please select a template first",
        "tm_hint_builtin_noedit": "Built-in templates cannot be edited. Duplicate to a custom template first.",
        "tm_hint_builtin_nodelete": "Built-in templates cannot be deleted",
        "tm_confirm_delete": "Confirm Delete",
        "tm_confirm_delete_msg": 'Delete custom template "{}"?',
        "tm_import_title": "Select template to import",
        "tm_filetype_template": "Template Files",
        "tm_filetype_yaml": "YAML Files",
        "tm_filetype_json": "JSON Files",
        "tm_export_title": "Export Template",
        "tm_err_load": "Cannot load template: {}",
        "tm_exists_title": "Template Exists",
        "tm_exists_overwrite": "Custom template already exists. Overwrite?",
        "tm_overwrite_ok": "Template overwritten",
        "tm_add_title": "New Custom Template",
        "tm_dup_title": "Duplicate Template",
        "tm_based_on": 'Based on "{}"',
        "tm_label_key": "Key",
        "tm_label_name": "Template Name",
        "tm_cancel": "Cancel",
        "tm_create": "Create",
        "tm_duplicate_btn": "Duplicate",
        "te_title_new": "New Template",
        "te_title_edit": "Edit Template",
        "te_err_load": "Cannot load editor:\n{}",
        "te_label_name": "Template Name",
        "te_label_desc": "Description",
        "te_tab_paper": " Paper & Page ",
        "te_tab_spacing": " Spacing & Paragraph ",
        "te_tab_headings": " Heading Styles ",
        "te_tab_misc": " Body · PageNum · Figures ",
        "te_paper": "Paper",
        "te_size": "Size",
        "te_orient": "Orientation",
        "te_margin_body": "Body Margins (cm)",
        "te_margin_cover": "Cover Margins (cm)",
        "te_bind": "Binding",
        "te_bind_extra": "Extra Left Margin (cm)",
        "te_top": "Top",
        "te_bottom": "Bottom",
        "te_left": "Left",
        "te_right": "Right",
        "te_ls": "Line Spacing",
        "te_ls_type": "Type",
        "te_ls_value": "Value",
        "te_ps": "Paragraph Spacing (pt)",
        "te_ps_before": "Before",
        "te_ps_after": "After",
        "te_indent": "First-line Indent",
        "te_indent_enabled": "Enable",
        "te_indent_cm": "Value (cm)",
        "te_indent_chars": "Chars",
        "te_h_title": "Abstract/Ack/TOC",
        "te_h1": "Heading 1",
        "te_h2": "Heading 2",
        "te_h3": "Heading 3",
        "te_h4": "Heading 4",
        "te_zh_font": "Chinese Font",
        "te_en_font": "English Font",
        "te_font_size": "Size",
        "te_bold": "Bold",
        "te_yes": "Yes",
        "te_no": "No",
        "te_align": "Align",
        "te_before": "Before (pt)",
        "te_after": "After (pt)",
        "te_page_break": "Page break before",
        "te_body": "Body Style",
        "te_fm": "Front-matter Page Number",
        "te_fm_style": "Style",
        "te_fm_pos": "Position",
        "te_fm_align": "Align",
        "te_bm": "Body Page Number",
        "te_bm_start": "Start Number",
        "te_fig": "Figure Format",
        "te_fig_num": "Numbering",
        "te_fig_caption": "Caption Below",
        "te_tab": "Table Format",
        "te_tab_caption": "Caption Above",
        "te_save": "Save",
        "te_err_name": "Validation Error",
        "te_err_name_msg": "Template name cannot be empty",
    },
}

_current = DEFAULT_LANG
_subscribers = []


def _load_lang():
    global _current
    try:
        if CONFIG_PATH.exists():
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line.startswith("language"):
                        val = line.split("=", 1)[1].strip().strip('"').strip("'")
                        if val in LANGS:
                            _current = val
                            return
    except Exception:
        pass


def _save_lang():
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            f.write('language = "{}"\n'.format(_current))
    except Exception:
        pass


_load_lang()


def get_lang():
    return _current


def set_lang(lang):
    global _current
    if lang not in LANGS:
        return
    if lang == _current:
        return
    _current = lang
    _save_lang()
    for cb in list(_subscribers):
        try:
            cb()
        except Exception:
            pass


def subscribe(cb):
    if cb not in _subscribers:
        _subscribers.append(cb)


def unsubscribe(cb):
    if cb in _subscribers:
        _subscribers.remove(cb)


def t(key, *args):
    s = _STRINGS.get(_current, _STRINGS[DEFAULT_LANG]).get(key, key)
    if args:
        try:
            return s.format(*args)
        except Exception:
            return s
    return s
