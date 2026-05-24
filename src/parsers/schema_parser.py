import re
from src.models import TableInfo, ColumnInfo, ForeignKeyInfo


def parse_schema(filepath: str) -> list:
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    tables = []
    table_blocks = re.split(r'CREATE TABLE IF NOT EXISTS\s+', content, flags=re.IGNORECASE)
    for block in table_blocks[1:]:
        table = _parse_table_block(block)
        if table:
            tables.append(table)
    return tables


def _parse_table_block(block: str) -> TableInfo:
    lines = block.split('\n')
    table_name = lines[0].strip().split('(')[0].strip()

    body_start = block.index('(') + 1
    idx = block.rfind(')')
    body = block[body_start:idx]

    columns = []
    foreign_keys = []

    statements = _split_sql_statements(body)

    for stmt in statements:
        stmt = stmt.strip()
        if not stmt:
            continue

        fk_match = re.match(
            r'FOREIGN\s+KEY\s*\((\w+)\)\s*REFERENCES\s*(\w+)\s*\((\w+)\)\s*(.*)',
            stmt, re.IGNORECASE
        )
        if fk_match:
            on_delete = ''
            on_delete_match = re.search(r'ON\s+DELETE\s+(\w+)', fk_match.group(4), re.IGNORECASE)
            if on_delete_match:
                on_delete = on_delete_match.group(1)
            foreign_keys.append(ForeignKeyInfo(
                column=fk_match.group(1),
                ref_table=fk_match.group(2),
                ref_column=fk_match.group(3),
                on_delete=on_delete
            ))
            continue

        col = _parse_column_def(stmt)
        if col:
            columns.append(col)

    return TableInfo(name=table_name, columns=columns, foreign_keys=foreign_keys)


def _parse_column_def(stmt: str) -> ColumnInfo:
    match = re.match(r'(\w+)\s+(\w+)(\([\d,]+\))?\s*(.*)', stmt, re.IGNORECASE | re.DOTALL)
    if not match:
        return None

    name = match.group(1)
    type_name = match.group(2).upper()
    type_len_str = match.group(3)
    options = match.group(4)

    length = None
    if type_len_str:
        nums = type_len_str.strip('()')
        length = int(nums.split(',')[0])

    sql_type = type_name
    if type_len_str:
        sql_type = f"{type_name}{type_len_str}"

    nullable = 'NOT NULL' not in options.upper()
    key_type = _extract_key_type(options)

    default_value = ''
    default_match = re.search(r"DEFAULT\s+('[^']*'|\S+)", options, re.IGNORECASE)
    if default_match:
        default_value = default_match.group(1)
    if re.search(r'DEFAULT\s+CURRENT_TIMESTAMP', options, re.IGNORECASE):
        default_value = 'CURRENT_TIMESTAMP'

    return ColumnInfo(
        name=name,
        sql_type=sql_type,
        type_name=type_name,
        length=length,
        nullable=nullable,
        key_type=key_type,
        default_value=default_value,
    )


def _extract_key_type(options: str) -> str:
    keys = []
    upper = options.upper()
    if 'AUTO_INCREMENT' in upper or 'PRIMARY KEY' in upper:
        keys.append('PK')
    if 'UNIQUE' in upper:
        keys.append('UNIQUE')
    return ', '.join(keys)


def _split_sql_statements(body: str) -> list:
    result = []
    depth = 0
    current = []
    for ch in body:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        elif ch == ',' and depth == 0:
            result.append(''.join(current))
            current = []
            continue
        current.append(ch)
    if current:
        result.append(''.join(current))
    return result
