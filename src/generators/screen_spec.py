from src.generators.styles import (
    apply_title_style, apply_header_style, apply_body_style,
    write_header_row, write_data_row
)


def generate_screen_specs(wb, screens, styles):
    for screen in screens:
        sheet_name = f"{screen.id}_{screen.name}"[:31]
        ws = wb.create_sheet(title=sheet_name)

        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 16
        ws.column_dimensions['F'].width = 40

        row = 1
        cell = ws.cell(row=row, column=1, value="画面仕様書")
        apply_title_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1

        info_items = [
            ("画面ID", screen.id), ("画面名", screen.name), ("URL", screen.url),
            ("JSPファイル", screen.jsp_file), ("Controller", screen.controller),
        ]
        for label, value in info_items:
            cell = ws.cell(row=row, column=1, value=label)
            apply_header_style(cell, styles)
            cell = ws.cell(row=row, column=2, value=value)
            apply_body_style(cell, styles)
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
            row += 1

        if screen.is_mock:
            cell = ws.cell(row=row, column=1, value="※ サンプル実装 / 本番未実装")
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            row += 1

        row += 1
        cell = ws.cell(row=row, column=1, value="画面レイアウト")
        apply_header_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1
        cell = ws.cell(row=row, column=1, value=screen.layout_description)
        apply_body_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 2

        if screen.table_columns:
            _write_section(ws, row, "表示項目一覧", styles)
            row += 1
            write_header_row(ws, row, ["No", "項目キー", "表示ラベル", "幅", "データ型", ""], styles)
            row += 1
            for i, col in enumerate(screen.table_columns):
                write_data_row(ws, row, [i+1, col.key, col.label, col.width, col.data_type, ""], styles)
                row += 1
            row += 1

        if screen.fields:
            _write_section(ws, row, "入力項目定義", styles)
            row += 1
            write_header_row(ws, row, ["No", "項目名", "型", "桁数", "必須", "備考"], styles)
            row += 1
            for i, f in enumerate(screen.fields):
                req_str = "○" if f.required else ""
                write_data_row(ws, row, [i+1, f.name, f.field_type, f.length or "-", req_str, f.description], styles)
                row += 1
            row += 1

        if screen.buttons:
            _write_section(ws, row, "操作ボタン一覧", styles)
            row += 1
            write_header_row(ws, row, ["No", "ボタン名", "アクション", "備考", "", ""], styles)
            row += 1
            for i, b in enumerate(screen.buttons):
                write_data_row(ws, row, [i+1, b.name, b.action, b.description or "", "", ""], styles)
                row += 1
            row += 1

        if screen.validations:
            _write_section(ws, row, "バリデーションルール", styles)
            row += 1
            write_header_row(ws, row, ["No", "対象項目", "ルール", "メッセージ", "", ""], styles)
            row += 1
            for i, v in enumerate(screen.validations):
                write_data_row(ws, row, [i+1, v.field, v.rule, v.message, "", ""], styles)
                row += 1


def _write_section(ws, row, title, styles):
    cell = ws.cell(row=row, column=1, value=title)
    apply_header_style(cell, styles)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
