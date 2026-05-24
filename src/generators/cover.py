from datetime import date
from src.generators.styles import apply_title_style, write_header_row, write_data_row, apply_body_style


def generate_cover(ws, sys_info, styles):
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 18

    cell = ws['A1']
    cell.value = "システム詳細設計書"
    apply_title_style(cell, styles)
    ws.merge_cells('A1:D1')

    row = 3
    info_items = [
        ("システム名", sys_info.name),
        ("バージョン", sys_info.version),
        ("パッケージング", sys_info.packaging),
        ("作成日", date.today().strftime('%Y-%m-%d')),
    ]
    for label, value in info_items:
        ws.cell(row=row, column=1, value=label).font = styles['header_font']
        cell = ws.cell(row=row, column=2, value=value)
        apply_body_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        row += 1

    headers = ["版数", "更新日", "更新内容", "作成者"]
    write_header_row(ws, row, headers, styles)
    row += 1

    records = [
        ("1.0", date.today().strftime('%Y-%m-%d'), "初版作成", "自動生成"),
    ]
    for rec in records:
        write_data_row(ws, row, rec, styles)
        row += 1
