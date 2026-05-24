from openpyxl import Workbook
from src.models import TableInfo
from src.generators.er_diagram import generate_er_diagram, build_relationships
from src.generators.styles import create_styles


def test_generate_er_diagram():
    wb = Workbook()
    ws = wb.active
    tables = [TableInfo(name="departments", logical_name="部署マスタ")]
    relationships = [("departments", "departments", "1", "N", "parent_id")]
    styles = create_styles()
    generate_er_diagram(ws, tables, relationships, styles)
    assert ws['A1'].value == "ER図（エンティティ関連図）"
