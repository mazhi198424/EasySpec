import tempfile
import os
from src.parsers.entity_parser import parse_entity


def test_parse_entity_extracts_table_and_columns(sample_entity_content):
    with tempfile.NamedTemporaryFile(mode='w', suffix='.java', delete=False, encoding='utf-8') as f:
        f.write(sample_entity_content)
        tmp_path = f.name

    try:
        result = parse_entity(tmp_path)
        assert result['table_name'] == "employees"
        assert len(result['columns']) == 4

        # id column
        id_col = result['columns'][0]
        assert id_col['name'] == 'id'
        assert id_col['java_type'] == 'Long'
        assert id_col['is_id'] is True

        # employee_no column with @Column annotations
        no_col = result['columns'][1]
        assert no_col['name'] == 'employee_no'
        assert no_col['java_type'] == 'String'
        assert no_col['nullable'] is False
        assert no_col['unique'] is True
        assert no_col['length'] == 20

        # name column
        name_col = result['columns'][2]
        assert name_col['name'] == 'name'
        assert name_col['nullable'] is False
        assert name_col['length'] == 100

        # department FK
        dept_col = result['columns'][3]
        assert dept_col['name'] == 'department_id'
        assert dept_col['is_fk'] is True
        assert dept_col['fk_entity'] == 'Department'

        assert len(result['foreign_keys']) == 1
        assert result['foreign_keys'][0]['column'] == 'department_id'
        assert result['foreign_keys'][0]['ref_table'] == 'departments'
        assert result['foreign_keys'][0]['ref_column'] == 'id'
    finally:
        os.unlink(tmp_path)
