from src.generators.styles import apply_title_style, write_header_row, write_data_row, apply_header_style, apply_body_style


def generate_table_def(ws, tables, styles):
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 16
    ws.column_dimensions['H'].width = 30

    headers = ["No", "論理名", "物理名", "型", "桁数", "NULL", "KEY", "説明"]
    row = 1

    for ti, table in enumerate(tables):
        if ti > 0:
            row += 1

        title = f"テーブル: {table.name} ({table.logical_name})"
        cell = ws.cell(row=row, column=1, value=title)
        apply_title_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        row += 1

        write_header_row(ws, row, headers, styles)
        row += 1

        for ci, col in enumerate(table.columns):
            null_label = "" if col.nullable else "NOT NULL"
            data = [ci + 1, col.logical_name or col.name, col.name, col.sql_type,
                    col.length if col.length else "", null_label, col.key_type, col.description or ""]
            write_data_row(ws, row, data, styles)
            row += 1

        if table.foreign_keys:
            cell = ws.cell(row=row, column=1, value="外部キー関連:")
            apply_header_style(cell, styles)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            row += 1

            fk_headers = ["", "FKカラム", "参照テーブル", "参照カラム", "削除ルール", "", "", ""]
            write_header_row(ws, row, fk_headers, styles)
            row += 1

            for fk in table.foreign_keys:
                fk_data = ["", fk.column, fk.ref_table, fk.ref_column, fk.on_delete or "RESTRICT", "", "", ""]
                write_data_row(ws, row, fk_data, styles)
                row += 1
