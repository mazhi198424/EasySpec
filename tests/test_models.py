# tests/test_models.py
from src.models import (
    SystemInfo, ColumnInfo, ForeignKeyInfo, TableInfo,
    TableColumnDef, ScreenField, ScreenButton, ValidationRule,
    ScreenInfo, ApiParam, ApiEndpoint
)


def test_system_info_creation():
    sys = SystemInfo(name="TestApp", version="1.0.0", packaging="war")
    assert sys.name == "TestApp"
    assert sys.version == "1.0.0"
    assert sys.packaging == "war"


def test_column_info_defaults():
    col = ColumnInfo(name="id", sql_type="BIGINT")
    assert col.name == "id"
    assert col.nullable is True
    assert col.key_type == ""
    assert col.length is None


def test_table_info_with_columns():
    cols = [
        ColumnInfo(name="id", sql_type="BIGINT", key_type="PK"),
        ColumnInfo(name="name", sql_type="VARCHAR", length=100, nullable=False),
    ]
    fks = [ForeignKeyInfo(column="dept_id", ref_table="departments", ref_column="id")]
    table = TableInfo(name="employees", logical_name="社員", columns=cols, foreign_keys=fks)
    assert len(table.columns) == 2
    assert table.columns[0].key_type == "PK"
    assert table.foreign_keys[0].ref_table == "departments"


def test_api_endpoint_creation():
    ep = ApiEndpoint(
        method="GET",
        url="/employee/api/search",
        controller="EmployeeController",
        request_params=[
            ApiParam(name="name", location="query", required=False),
            ApiParam(name="page", location="query", required=False, param_type="int"),
        ],
        response_type="Map<String, Object>",
        description="社員検索"
    )
    assert ep.method == "GET"
    assert len(ep.request_params) == 2
    assert ep.request_params[0].location == "query"


def test_screen_info_creation():
    screen = ScreenInfo(
        id="SCR-003",
        name="社員管理",
        url="/employee/page",
        jsp_file="fragments/employee-list.jsp",
        controller="EmployeeController",
        fields=[ScreenField(name="emp-search-name", field_type="text", description="氏名検索")],
        buttons=[ScreenButton(name="emp-search-btn", action="検索")],
        table_columns=[
            TableColumnDef(key="employeeNo", label="社員番号", width="90px"),
            TableColumnDef(key="name", label="氏名", width="120px"),
        ],
        validations=[ValidationRule(field="emp-search-name", rule="任意入力", message="")]
    )
    assert screen.id == "SCR-003"
    assert len(screen.table_columns) == 2
    assert screen.fields[0].name == "emp-search-name"
