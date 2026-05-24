import re
import os


def parse_entity(filepath: str) -> dict:
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    class_name = os.path.splitext(os.path.basename(filepath))[0]
    table_name = _extract_table_name(content, class_name)

    columns = []
    foreign_keys = []

    field_blocks = _split_fields(content)

    for block in field_blocks:
        col_info = _parse_field_block(block, table_name)
        if col_info['is_fk']:
            foreign_keys.append({
                'column': col_info['name'],
                'ref_table': col_info.get('fk_table') or '',
                'ref_column': col_info.get('fk_column') or 'id',
                'on_delete': col_info.get('on_delete') or '',
            })
        columns.append(col_info)

    return {
        'table_name': table_name,
        'entity_class': class_name,
        'columns': columns,
        'foreign_keys': foreign_keys,
    }


def _extract_table_name(content: str, class_name: str) -> str:
    m = re.search(r'@Table\s*\(\s*name\s*=\s*"(\w+)"', content)
    if m:
        return m.group(1)
    return class_name.lower() + 's'


def _split_fields(content: str) -> list:
    pattern = r'((?:@\w+[^;]*\s*)*)\s*private\s+(\w+(?:<\w+>)?)\s+(\w+)\s*;'
    matches = re.findall(pattern, content)
    result = []
    for annotations_raw, java_type, field_name in matches:
        result.append({
            'annotations': annotations_raw,
            'java_type': java_type,
            'field_name': field_name,
        })
    return result


def _parse_field_block(field_data: dict, table_name: str) -> dict:
    annotations = field_data['annotations']
    java_type = field_data['java_type']
    field_name = field_data['field_name']

    is_id = '@Id' in annotations
    is_fk = '@ManyToOne' in annotations or '@OneToMany' in annotations

    col_info = {
        'name': field_name,
        'java_type': java_type,
        'is_id': is_id,
        'is_fk': is_fk,
        'nullable': True,
        'unique': False,
        'length': None,
        'fk_entity': None,
        'fk_table': None,
        'fk_column': None,
        'on_delete': '',
    }

    # Extract @Column properties
    col_match = re.search(r'@Column\s*\(([^)]+)\)', annotations)
    if col_match:
        props = col_match.group(1)
        name_m = re.search(r'name\s*=\s*"(\w+)"', props)
        if name_m:
            col_info['name'] = name_m.group(1)

        if 'nullable' in props:
            null_m = re.search(r'nullable\s*=\s*(true|false)', props)
            if null_m:
                col_info['nullable'] = null_m.group(1) == 'true'

        if 'unique' in props:
            uniq_m = re.search(r'unique\s*=\s*(true|false)', props)
            if uniq_m:
                col_info['unique'] = uniq_m.group(1) == 'true'

        len_m = re.search(r'length\s*=\s*(\d+)', props)
        if len_m:
            col_info['length'] = int(len_m.group(1))

    # Extract @JoinColumn for FK
    join_match = re.search(r'@JoinColumn\s*\(([^)]+)\)', annotations)
    if join_match:
        props = join_match.group(1)
        name_m = re.search(r'name\s*=\s*"(\w+)"', props)
        if name_m:
            col_info['name'] = name_m.group(1)

    # FK entity
    if is_fk and '@ManyToOne' in annotations:
        col_info['fk_entity'] = java_type
        col_info['fk_table'] = java_type.lower() + 's'

    return col_info
