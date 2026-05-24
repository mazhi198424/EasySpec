# src/parsers/jsp_parser.py
import re
from src.models import ScreenField, ScreenButton, TableColumnDef


def parse_jsp(filepath: str) -> dict:
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    fields = _extract_input_fields(content)
    buttons = _extract_buttons(content)
    table_columns = _extract_table_columns(content)
    has_search = bool(re.search(r'search-bar|search-btn', content))
    has_table = bool(re.search(r'table-container|renderTable', content))
    layout_parts = _describe_layout(content)

    return {
        'fields': fields,
        'buttons': buttons,
        'table_columns': table_columns,
        'has_search': has_search,
        'has_table': has_table,
        'layout_description': ' / '.join(layout_parts) if layout_parts else '',
    }


def _extract_input_fields(content: str) -> list:
    fields = []
    # Match <input> elements with id attributes
    for m in re.finditer(
        r'<input\s+[^>]*?id\s*=\s*"(\w+(?:-\w+)*)"[^>]*?(?:type\s*=\s*"(\w+)")?[^>]*>',
        content
    ):
        input_id = m.group(1)
        input_type = m.group(2) or 'text'
        fields.append(ScreenField(name=input_id, field_type=input_type))

    # Match <select> elements
    for m in re.finditer(r'<select\s+[^>]*?id\s*=\s*"(\w+(?:-\w+)*)"', content):
        fields.append(ScreenField(name=m.group(1), field_type='select'))

    # Match <textarea> elements
    for m in re.finditer(r'<textarea\s+[^>]*?id\s*=\s*"(\w+(?:-\w+)*)"', content):
        fields.append(ScreenField(name=m.group(1), field_type='textarea'))

    return fields


def _extract_buttons(content: str) -> list:
    buttons = []
    for m in re.finditer(
        r'<button\s+[^>]*?id\s*=\s*"(\w+(?:-\w+)*)"[^>]*>(.*?)</button>',
        content, re.DOTALL
    ):
        label = re.sub(r'<[^>]+>', '', m.group(2)).strip()
        buttons.append(ScreenButton(name=m.group(1), action=label))
    return buttons


def _extract_table_columns(content: str) -> list:
    """Extract table column definitions from JS renderTable calls."""
    columns = []
    pattern = r"\{\s*key\s*:\s*'(\w+)'\s*,\s*label\s*:\s*'([^']+)'\s*(?:,\s*width\s*:\s*'([^']+)')?"
    for m in re.finditer(pattern, content):
        columns.append(TableColumnDef(
            key=m.group(1),
            label=m.group(2),
            width=m.group(3) or '',
        ))
    return columns


def _describe_layout(content: str) -> list:
    parts = []
    if 'search-bar' in content or 'search-btn' in content:
        parts.append('検索エリア')
    if 'table-container' in content or 'renderTable' in content:
        parts.append('一覧テーブル')
    if 'pagination' in content:
        parts.append('ページネーション')
    if 'modal' in content.lower() or 'showModal' in content:
        parts.append('モーダル編集')
    if 'tab-item' in content:
        parts.append('タブ切替')
    if 'tree-toggle' in content or 'tree-item' in content:
        parts.append('ツリー表示')
    if 'wizard' in content.lower() or 'wiz-step' in content:
        parts.append('ウィザード形式')
    return parts
