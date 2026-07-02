import yaml
import sys
from pathlib import Path
from typing import Dict, Any, List, Optional
from .debug_log import get_logger
from .template_manager import TemplateManager, BUILTIN_TEMPLATES

log = get_logger("rule_engine")

TEMPLATES = BUILTIN_TEMPLATES.copy()

def _resolve_rule_path(rule_file: str) -> Path:
    if getattr(sys, 'frozen', False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).parent.parent
    builtin = base / "rules" / rule_file
    if builtin.exists():
        return builtin
    custom = base / "rules" / "custom" / rule_file
    if custom.exists():
        return custom
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent / "rules" / "custom" / rule_file
    return base / "rules" / rule_file

class RuleEngine:
    def __init__(self, rule_file: str = "thesis_zh.yaml"):
        self.rules_path = _resolve_rule_path(rule_file)
        log.info("RuleEngine __init__: rule_file=%s, full_path=%s, exists=%s",
                 rule_file, self.rules_path, self.rules_path.exists())
        self.rules = self._load_rules()
    
    def _load_rules(self) -> Dict[str, Any]:
        if not self.rules_path.exists():
            log.error("規則檔不存在: %s", self.rules_path)
            return {}
        with open(self.rules_path, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)
        log.info("規則載入成功, keys=%s", list(data.keys()) if data else [])
        return data or {}
    
    def get(self, key: str, default=None):
        keys = key.split('.')
        value = self.rules
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        return value
    
    def get_page_margins(self, section_type: str = "body") -> tuple:
        section = self.rules.get('page_margins', {}).get(section_type,
                    self.rules.get('page_margins', {}).get('body', {}))
        return (section.get('top', 2.5), section.get('bottom', 2.5),
                section.get('left', 2.7), section.get('right', 2.7))
    
    def get_font_settings(self) -> Dict:
        return self.rules.get('fonts', {})
    
    def get_heading_style(self, level: int) -> Dict:
        styles = self.rules.get('heading_styles', {})
        names = ['heading1', 'heading2', 'heading3', 'heading4']
        if 0 <= level < len(names):
            return styles.get(names[level], {})
        return {}
    
    def get_line_spacing(self) -> Dict:
        return self.rules.get('line_spacing', {})
    
    def get_cover_settings(self) -> Dict:
        return self.rules.get('cover', {})
    
    def get_numbered_list_rules(self) -> Dict:
        return self.rules.get('numbered_list', {})
    
    def get_heading_numbering_conversion(self) -> Dict:
        return self.rules.get('heading_numbering_conversion', {})
    
    def get_figure_table_rules(self) -> Dict:
        return {
            'figures': self.rules.get('figures', {}),
            'tables': self.rules.get('tables', {}),
        }
    
    def get_template_name(self) -> str:
        return self.rules.get('template', {}).get('name', 'Unknown')
    
    @staticmethod
    def available_templates() -> List[Dict[str, str]]:
        return TemplateManager.get_all_templates()

    @staticmethod
    def get_template_info(key: str) -> dict:
        return TemplateManager.get_template_info(key)

    @staticmethod
    def create(template_key: str = 'zh') -> 'RuleEngine':
        info = TemplateManager.get_template_info(template_key)
        filename = BUILTIN_TEMPLATES.get(template_key, f"{template_key}.yaml")
        return RuleEngine(filename)
