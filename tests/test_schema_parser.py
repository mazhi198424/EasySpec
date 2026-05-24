import tempfile
import os
from src.parsers.schema_parser import parse_schema


def test_parse_schema_returns_table_list(sample_schema_content):
    with tempfile.NamedTemporaryFile(mode='w', suffix='.sql', delete=False, encoding='utf-8') as f:
        f.write(sample_schema_content)
        tmp_path = f.name

    try:
        tables = parse_schema(tmp_path)
        assert len(tables) == 2

        # departments table
        dept = tables[0]
        assert dept.name == "departments"
        assert len(dept.columns) == 6
        assert dept.columns[0].name == "id"
        assert dept.columns[0].sql_type == "BIGINT"
        assert dept.columns[0].key_type == "PK"

        # code column
        assert dept.columns[1].name == "code"
        assert dept.columns[1].type_name == "VARCHAR"
        assert dept.columns[1].length == 20
        assert dept.columns[1].nullable is False
        assert "UNIQUE" in dept.columns[1].key_type

        # parent_id FK
        assert dept.columns[3].name == "parent_id"
        assert len(dept.foreign_keys) == 1
        assert dept.foreign_keys[0].column == "parent_id"
        assert dept.foreign_keys[0].ref_table == "departments"

        # employees table
        emp = tables[1]
        assert emp.name == "employees"
        assert len(emp.columns) == 4
        assert emp.columns[0].name == "id"
        assert emp.columns[0].key_type == "PK"
        assert emp.foreign_keys[0].column == "department_id"
        assert emp.foreign_keys[0].ref_table == "departments"
    finally:
        os.unlink(tmp_path)


def test_column_with_default_value():
    sql = """CREATE TABLE IF NOT EXISTS test_table (
    id BIGINT AUTO_INCREMENT PRIMARY KEY,
    status VARCHAR(20) DEFAULT '下書き'
);"""
    with tempfile.NamedTemporaryFile(mode='w', suffix='.sql', delete=False, encoding='utf-8') as f:
        f.write(sql)
        tmp_path = f.name

    try:
        tables = parse_schema(tmp_path)
        assert tables[0].columns[1].default_value == "'下書き'"
    finally:
        os.unlink(tmp_path)
