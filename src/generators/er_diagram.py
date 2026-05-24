from src.generators.styles import apply_title_style, apply_header_style, apply_body_style


def generate_er_diagram(ws, tables, relationships, styles):
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 28
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 28
    ws.column_dimensions['F'].width = 20

    cell = ws['A1']
    cell.value = "ER図（エンティティ関連図）"
    apply_title_style(cell, styles)
    ws.merge_cells('A1:F1')

    row = 3
    ws.cell(row=row, column=1, value="凡例:").font = styles['header_font']
    row += 1
    ws.cell(row=row, column=1, value="■ = エンティティ（テーブル）").font = styles['body_font']
    row += 1
    ws.cell(row=row, column=1, value="→ = リレーション（外部キー）").font = styles['body_font']
    row += 2

    for table in tables:
        cell = ws.cell(row=row, column=1, value=f"■ {table.name}")
        apply_header_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        row += 1
        cell = ws.cell(row=row, column=1, value=table.logical_name)
        apply_body_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        row += 1
        for col in table.columns[:5]:
            key_mark = "PK" if "PK" in col.key_type else "  "
            cell = ws.cell(row=row, column=1, value=f"[{key_mark}] {col.name}")
            apply_body_style(cell, styles)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            row += 1
        row += 1

    row += 1
    cell = ws.cell(row=row, column=1, value="リレーション一覧")
    apply_header_style(cell, styles)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    row += 1

    rel_headers = ["親テーブル", "子テーブル", "多重度(親)", "多重度(子)", "FKカラム", "備考"]
    for i, h in enumerate(rel_headers):
        cell = ws.cell(row=row, column=i+1, value=h)
        apply_header_style(cell, styles)
    row += 1

    for parent, child, card_p, card_c, fk_col in relationships:
        data = [parent, child, card_p, card_c, fk_col, f"{child}.{fk_col} → {parent}.id"]
        for i, val in enumerate(data):
            cell = ws.cell(row=row, column=i+1, value=val)
            apply_body_style(cell, styles)
        row += 1


def build_relationships(tables) -> list:
    relationships = []
    for table in tables:
        for fk in table.foreign_keys:
            relationships.append((fk.ref_table, table.name, "1", "N", fk.column))
    return relationships
