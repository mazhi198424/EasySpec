from src.generators.styles import apply_title_style, write_header_row, write_data_row


def generate_screen_list(ws, screens, styles):
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 35
    ws.column_dimensions['E'].width = 50

    cell = ws['A1']
    cell.value = "画面一覧"
    apply_title_style(cell, styles)
    ws.merge_cells('A1:E1')

    headers = ["画面ID", "画面名", "URL", "JSPファイル", "機能概要"]
    row = 2
    write_header_row(ws, row, headers, styles)
    row += 1

    for screen in screens:
        data = [screen.id, screen.name, screen.url, screen.jsp_file,
                screen.layout_description]
        write_data_row(ws, row, data, styles)
        row += 1
