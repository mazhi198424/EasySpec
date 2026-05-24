# tests/test_project_reader.py
from src.project_reader import ProjectReader


def test_project_reader_with_visiondemo(visiondemo_path):
    reader = ProjectReader(visiondemo_path)

    # System info
    sys_info = reader.read_system_info()
    assert sys_info.name == "VisionDemo"
    assert sys_info.version == "1.0.0"

    # Tables
    tables = reader.read_tables()
    table_names = [t.name for t in tables]
    assert "departments" in table_names
    assert "employees" in table_names
    assert "expenses" in table_names
    assert "expense_details" in table_names
    assert "orders" in table_names
    assert "order_items" in table_names
    assert len(tables) == 6

    # API endpoints
    endpoints, _ = reader.read_api_endpoints()
    assert len(endpoints) > 0
    methods = set(ep.method for ep in endpoints)
    assert "GET" in methods
    assert "POST" in methods
    assert "DELETE" in methods

    # Screens
    screens = reader.read_screens()
    assert len(screens) >= 8
    screen_names = [s.name for s in screens]
    assert "組織管理" in screen_names
    assert "社員管理" in screen_names
