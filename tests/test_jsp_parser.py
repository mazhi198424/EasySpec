# tests/test_jsp_parser.py
import tempfile
import os
from src.parsers.jsp_parser import parse_jsp


def test_parse_jsp_extracts_fields_buttons_columns(sample_jsp_content):
    with tempfile.NamedTemporaryFile(mode='w', suffix='.jsp', delete=False, encoding='utf-8') as f:
        f.write(sample_jsp_content)
        tmp_path = f.name

    try:
        result = parse_jsp(tmp_path)

        # Fields (input elements)
        assert len(result['fields']) >= 1
        search_name = [f for f in result['fields'] if f.name == 'emp-search-name'][0]
        assert search_name.field_type == 'text'

        # Buttons
        assert len(result['buttons']) >= 2
        btn_names = [b.name for b in result['buttons']]
        assert 'emp-search-btn' in btn_names
        assert 'emp-reset-btn' in btn_names

        # Table columns (from renderTable)
        assert len(result['table_columns']) == 3
        assert result['table_columns'][0].key == 'employeeNo'
        assert result['table_columns'][0].label == '社員番号'
        assert result['table_columns'][1].key == 'name'
        assert result['table_columns'][1].label == '氏名'
        assert result['table_columns'][1].width == '120px'

        # Layout
        assert result['has_search'] is True
        assert result['has_table'] is True
    finally:
        os.unlink(tmp_path)
