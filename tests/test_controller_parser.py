# tests/test_controller_parser.py
import tempfile
import os
from src.parsers.controller_parser import parse_controller


def test_parse_controller_extracts_endpoints_and_pages(sample_controller_content):
    with tempfile.NamedTemporaryFile(mode='w', suffix='.java', delete=False, encoding='utf-8') as f:
        f.write(sample_controller_content)
        tmp_path = f.name

    try:
        endpoints, jsp_map = parse_controller(tmp_path)

        # 4 endpoints: page, api/search, api/save, api/{id}
        assert len(endpoints) == 4

        # GET /employee/page → JSP mapping
        page_ep = [e for e in endpoints if e.url == "/employee/page"][0]
        assert page_ep.method == "GET"
        assert page_ep.response_type == "String"
        assert jsp_map["/employee/page"] == "fragments/employee-list"

        # GET /employee/api/search
        search_ep = [e for e in endpoints if "/api/search" in e.url][0]
        assert search_ep.method == "GET"
        assert len(search_ep.request_params) == 4
        assert search_ep.request_params[0].name == "name"
        assert search_ep.request_params[0].location == "query"
        assert search_ep.request_params[0].required is False
        assert search_ep.request_params[0].param_type == "String"

        # POST /employee/api/save
        save_ep = [e for e in endpoints if "/api/save" in e.url][0]
        assert save_ep.method == "POST"
        has_body = any(p.location == "body" for p in save_ep.request_params)
        assert has_body is True

        # DELETE /employee/api/{id}
        del_ep = [e for e in endpoints if e.method == "DELETE"][0]
        assert del_ep.url == "/employee/api/{id}"
        has_path = any(p.location == "path" for p in del_ep.request_params)
        assert has_path is True

    finally:
        os.unlink(tmp_path)
