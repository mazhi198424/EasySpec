from src.generators.styles import apply_title_style, apply_header_style, apply_body_style, write_header_row, write_data_row


def generate_business_flow(ws, styles):
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 30
    ws.column_dimensions['F'].width = 20

    cell = ws['A1']
    cell.value = "業務フロー"
    apply_title_style(cell, styles)
    ws.merge_cells('A1:F1')

    row = 3

    # Expense flow
    _write_flow_section(ws, row, "経費精算フロー", styles)
    row += 1
    expense_states = [
        ("1", "下書き", "経費明細を入力中の状態。保存可能。", "申請者が作成", "経費精算登録画面"),
        ("2", "申請中", "承認者に申請を行った状態。編集不可。", "申請ボタン押下", "ExpenseController.submit()"),
        ("3", "承認済", "承認者が承認した状態。処理完了。", "承認者が承認", "（本実装では画面未実装）"),
        ("4", "差戻し", "承認者が差し戻した状態。再編集可能。", "承認者が差戻し", "（本実装では画面未実装）"),
    ]
    _write_state_table(ws, row, expense_states, styles)
    row += len(expense_states) + 2

    transitions = [
        "下書き → 申請中 : 申請ボタン押下（ExpenseController.submit）",
        "申請中 → 承認済 : 承認処理（未実装）",
        "申請中 → 差戻し : 差戻し処理（未実装）",
        "差戻し → 下書き : 再編集（未実装）",
    ]
    for t in transitions:
        cell = ws.cell(row=row, column=1, value=f"  {t}")
        apply_body_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1

    row += 2

    # Order flow
    _write_flow_section(ws, row, "発注フロー", styles)
    row += 1
    order_states = [
        ("1", "新規", "発注の初期状態。", "発注作成時", "OrderController.save()"),
        ("2", "発注済", "仕入先に発注済み。キャンセル可能。", "発注処理実行", "（未実装）"),
        ("3", "納品済", "商品が納品された状態。処理完了。", "納品確認", "（未実装）"),
        ("4", "キャンセル", "発注がキャンセルされた状態。", "キャンセル操作", "OrderController.cancel()"),
    ]
    _write_state_table(ws, row, order_states, styles)
    row += len(order_states) + 2

    order_transitions = [
        "新規 → 発注済 : 発注処理（未実装）",
        "発注済 → 納品済 : 納品確認（未実装）",
        "新規 → キャンセル : キャンセル操作",
        "発注済 → キャンセル : キャンセル操作（OrderController.cancel）",
    ]
    for t in order_transitions:
        cell = ws.cell(row=row, column=1, value=f"  {t}")
        apply_body_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1


def _write_flow_section(ws, row, title, styles):
    cell = ws.cell(row=row, column=1, value=title)
    apply_header_style(cell, styles)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)


def _write_state_table(ws, row, states, styles):
    headers = ["No", "状態名", "説明", "トリガー", "実装箇所"]
    write_header_row(ws, row, headers, styles)
    row += 1
    for s in states:
        write_data_row(ws, row, list(s), styles)
        row += 1
