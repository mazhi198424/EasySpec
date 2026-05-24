from src.generators.styles import apply_title_style, write_header_row, write_data_row


def generate_api_spec(ws, endpoints, styles):
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 40
    ws.column_dimensions['E'].width = 40
    ws.column_dimensions['F'].width = 30
    ws.column_dimensions['G'].width = 20

    cell = ws['A1']
    cell.value = "API仕様書"
    apply_title_style(cell, styles)
    ws.merge_cells('A1:G1')

    headers = ["No", "機能", "メソッド", "URL", "リクエストパラメータ", "レスポンス", "Controller"]
    row = 2
    write_header_row(ws, row, headers, styles)
    row += 1

    for i, ep in enumerate(endpoints):
        params_str = _format_params(ep.request_params)
        desc = ep.description or _infer_description(ep)
        data = [i + 1, desc, ep.method, ep.url, params_str, ep.response_type, ep.controller]
        write_data_row(ws, row, data, styles)
        row += 1


def _format_params(params) -> str:
    if not params:
        return "-"
    lines = []
    for p in params:
        req = "必須" if p.required else "任意"
        lines.append(f"{p.name} ({p.param_type}, {req}, {p.location})")
    return "\n".join(lines)


def _infer_description(ep) -> str:
    url_lower = ep.url.lower()
    method = ep.method
    if "search" in url_lower or "tree" in url_lower:
        return "検索・一覧取得"
    if "save" in url_lower or "add" in url_lower:
        return "新規登録・更新"
    if method == "DELETE":
        return "削除"
    if "submit" in url_lower:
        return "申請"
    if "cancel" in url_lower:
        return "キャンセル"
    if "copy" in url_lower:
        return "複製"
    if "move" in url_lower:
        return "移動"
    if "page" in url_lower:
        return "画面表示"
    if "param" in url_lower:
        return "パラメータ操作"
    if "dict" in url_lower:
        return "辞書操作"
    if "data" in url_lower:
        return "データ取得"
    return ""
