import yaml
import json
import sys
import shutil
from pathlib import Path
from typing import Dict, Any, List, Optional
from .debug_log import get_logger

log = get_logger("template_manager")

BUILTIN_TEMPLATES: Dict[str, str] = {
    'zh': 'thesis_zh.yaml',
}

class TemplateExistsError(Exception):
    pass

class TemplateNotFoundError(Exception):
    pass

class BuiltinTemplateError(Exception):
    pass

def _builtin_dir() -> Path:
    if getattr(sys, 'frozen', False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).parent.parent
    return base / "rules"

def _custom_dir() -> Path:
    if getattr(sys, 'frozen', False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).parent.parent
    d = base / "rules" / "custom"
    d.mkdir(parents=True, exist_ok=True)
    return d

def _load_any(path: Path) -> dict:
    ext = path.suffix.lower()
    with open(path, 'r', encoding='utf-8') as f:
        if ext == '.json':
            return json.load(f)
        return yaml.safe_load(f) or {}

def _save_any(path: Path, data: dict):
    ext = path.suffix.lower()
    with open(path, 'w', encoding='utf-8') as f:
        if ext == '.json':
            json.dump(data, f, ensure_ascii=False, indent=2)
        else:
            yaml.dump(data, f, allow_unicode=True, sort_keys=False, indent=2)

class TemplateManager:
    @staticmethod
    def _key_to_filename(key: str) -> str:
        return f"{key}.yaml"

    @staticmethod
    def _filename_to_key(filename: str) -> str:
        return Path(filename).stem

    @staticmethod
    def _is_builtin(key: str) -> bool:
        return key in BUILTIN_TEMPLATES

    @staticmethod
    def get_all_templates() -> List[Dict[str, str]]:
        result = []
        for key, filename in BUILTIN_TEMPLATES.items():
            path = _builtin_dir() / filename
            if path.exists():
                data = _load_any(path)
                name = data.get('template', {}).get('name', filename)
                desc = data.get('template', {}).get('description', '')
                result.append({'key': key, 'name': name, 'description': desc, 'file': filename, 'builtin': True})
        custom_dir = _custom_dir()
        if custom_dir.exists():
            for f in sorted(custom_dir.glob("*.yaml")):
                key = TemplateManager._filename_to_key(f.name)
                data = _load_any(f)
                name = data.get('template', {}).get('name', key)
                desc = data.get('template', {}).get('description', '')
                result.append({'key': key, 'name': name, 'description': desc, 'file': f.name, 'builtin': False})
        return result

    @staticmethod
    def load_template_data(key: str) -> dict:
        if key in BUILTIN_TEMPLATES:
            path = _builtin_dir() / BUILTIN_TEMPLATES[key]
        else:
            path = _custom_dir() / TemplateManager._key_to_filename(key)
        if not path.exists():
            raise TemplateNotFoundError(f"模板 '{key}' 不存在: {path}")
        return _load_any(path)

    @staticmethod
    def get_template_path(key: str) -> Path:
        if key in BUILTIN_TEMPLATES:
            return _builtin_dir() / BUILTIN_TEMPLATES[key]
        custom_path = _custom_dir() / TemplateManager._key_to_filename(key)
        if custom_path.exists():
            return custom_path
        builtin_path = _builtin_dir() / TemplateManager._key_to_filename(key)
        if builtin_path.exists():
            return builtin_path
        raise TemplateNotFoundError(f"模板 '{key}' 不存在")

    @staticmethod
    def create_template(key: str, name: str, description: str = "", base_on: Optional[str] = None) -> str:
        if key in BUILTIN_TEMPLATES:
            raise TemplateExistsError(f"不得覆蓋內建模板 '{key}'")
        custom_path = _custom_dir() / TemplateManager._key_to_filename(key)
        if custom_path.exists():
            raise TemplateExistsError(f"自訂模板 '{key}' 已存在")

        if base_on:
            data = TemplateManager.load_template_data(base_on)
        else:
            data = TemplateManager.load_template_data('zh')

        data['template'] = {'name': name, 'description': description}
        _save_any(custom_path, data)
        log.info("建立自訂模板: key=%s, name=%s, base_on=%s", key, name, base_on)
        return key

    @staticmethod
    def save_template(key: str, data: dict):
        if key in BUILTIN_TEMPLATES:
            raise BuiltinTemplateError(f"不得修改內建模板 '{key}'")
        custom_path = _custom_dir() / TemplateManager._key_to_filename(key)
        if not custom_path.exists():
            raise TemplateNotFoundError(f"自訂模板 '{key}' 不存在")
        _save_any(custom_path, data)
        log.info("已儲存自訂模板: key=%s", key)

    @staticmethod
    def delete_template(key: str):
        if key in BUILTIN_TEMPLATES:
            raise BuiltinTemplateError(f"不得刪除內建模板 '{key}'")
        custom_path = _custom_dir() / TemplateManager._key_to_filename(key)
        if not custom_path.exists():
            raise TemplateNotFoundError(f"自訂模板 '{key}' 不存在")
        custom_path.unlink()
        log.info("已刪除自訂模板: key=%s", key)

    @staticmethod
    def duplicate_template(source_key: str, new_key: str, new_name: str) -> str:
        if new_key in BUILTIN_TEMPLATES:
            raise TemplateExistsError(f"不得覆蓋內建模板 '{new_key}'")
        custom_path = _custom_dir() / TemplateManager._key_to_filename(new_key)
        if custom_path.exists():
            raise TemplateExistsError(f"自訂模板 '{new_key}' 已存在")
        data = TemplateManager.load_template_data(source_key)
        data['template']['name'] = new_name
        data['template']['description'] = data.get('template', {}).get('description', '') + f" (從 {source_key} 複製)"
        _save_any(custom_path, data)
        return new_key

    @staticmethod
    def import_template(filepath: str, key: Optional[str] = None) -> str:
        src = Path(filepath)
        if not src.exists():
            raise FileNotFoundError(f"檔案不存在: {filepath}")
        data = _load_any(src)
        actual_key = key or src.stem
        if actual_key in BUILTIN_TEMPLATES:
            raise TemplateExistsError(f"不得覆蓋內建模板 '{actual_key}'")
        dst = _custom_dir() / TemplateManager._key_to_filename(actual_key)
        if dst.exists():
            raise TemplateExistsError(f"自訂模板 '{actual_key}' 已存在")
        _save_any(dst, data)
        log.info("已匯入模板: %s → %s", filepath, dst)
        return actual_key

    @staticmethod
    def export_template(key: str, filepath: str):
        data = TemplateManager.load_template_data(key)
        dst = Path(filepath)
        _save_any(dst, data)
        log.info("已匯出模板: key=%s → %s", key, filepath)

    @staticmethod
    def get_template_info(key: str) -> dict:
        data = TemplateManager.load_template_data(key)
        tpl = data.get('template', {})
        return {
            'key': key,
            'name': tpl.get('name', key),
            'description': tpl.get('description', ''),
            'builtin': key in BUILTIN_TEMPLATES,
        }
