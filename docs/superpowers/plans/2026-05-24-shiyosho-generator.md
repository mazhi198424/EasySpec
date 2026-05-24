# VisionDemo 式様書 自動生成ツール 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** VisionDemo プロジェクトのソースコードを静的解析し、日本語の詳細設計書（式様書）Excel ファイルを自動生成する Python ツールを TDD で実装する。

**Architecture:** 3層構成 — (1) `parsers/` が各ソースファイル（pom.xml, schema.sql, Entity.java, Controller.java, JSP）を解析してデータ構造に変換、(2) `generators/` がそのデータ構造から openpyxl で Excel シートを生成、(3) `project_reader.py` が全パーサーを統括し `generate_shiyosho.py` がエントリポイントとなる。

**Tech Stack:** Python 3, openpyxl, pytest, xml.etree.ElementTree (stdlib)

---

## ファイル構成

```
EasySpec/
├── generate_shiyosho.py              # エントリポイント
├── requirements.txt                  # openpyxl==3.1.5
├── src/
│   ├── __init__.py
│   ├── models.py                     # 全データクラス
│   ├── parsers/
│   │   ├── __init__.py
│   │   ├── pom_parser.py             # pom.xml → SystemInfo
│   │   ├── schema_parser.py          # schema.sql → list[TableInfo]
│   │   ├── entity_parser.py          # Entity.java → アノテーション情報
│   │   ├── controller_parser.py      # Controller.java → list[ApiEndpoint] + URLマップ
│   │   └── jsp_parser.py             # JSP → ScreenInfo（部分）
│   ├── generators/
│   │   ├── __init__.py
│   │   ├── styles.py                 # Excel スタイル定数
│   │   ├── cover.py                  # 表紙シート
│   │   ├── screen_list.py            # 画面一覧シート
│   │   ├── table_def.py              # テーブル定義シート
│   │   ├── er_diagram.py             # ER図シート
│   │   ├── screen_spec.py            # 画面仕様シート
│   │   ├── api_spec.py               # API仕様書シート
│   │   └── business_flow.py          # 業務フローシート
│   └── project_reader.py             # 全パーサー統括（オーケストレータ）
└── tests/
    ├── __init__.py
    ├── conftest.py                    # 共通フィクスチャ
    ├── test_pom_parser.py
    ├── test_schema_parser.py
    ├── test_entity_parser.py
    ├── test_controller_parser.py
    ├── test_jsp_parser.py
    ├── test_project_reader.py
    ├── test_styles.py
    ├── test_cover.py
    ├── test_screen_list.py
    ├── test_table_def.py
    ├── test_er_diagram.py
    ├── test_screen_spec.py
    ├── test_api_spec.py
    ├── test_business_flow.py
    └── test_integration.py
```

---

### Task 1: プロジェクト初期化

**Files:**
- Create: `requirements.txt`
- Create: `src/__init__.py`
- Create: `src/parsers/__init__.py`
- Create: `src/generators/__init__.py`
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`

- [ ] **Step 1: ディレクトリ構造を作成**

```bash
mkdir -p src/parsers src/generators tests
touch src/__init__.py src/parsers/__init__.py src/generators/__init__.py tests/__init__.py
```

- [ ] **Step 2: requirements.txt を作成**

```
openpyxl==3.1.5
pytest==8.3.4
```

- [ ] **Step 3: tests/conftest.py を作成（共通フィクスチャ）**

```python
import pytest
import os

@pytest.fixture
def visiondemo_path():
    """VisionDemo プロジェクトの絶対パス"""
    return os.path.join(os.path.dirname(__file__), '..', '..', '..', 'VisionDemo')

@pytest.fixture
def sample_pom_content():
    return '''<?xml version="1.0" encoding="UTF-8"?>
<project xmlns="http://maven.apache.org/POM/4.0.0">
    <modelVersion>4.0.0</modelVersion>
    <parent>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-parent</artifactId>
        <version>3.4.5</version>
    </parent>
    <groupId>com.visiondemo</groupId>
    <artifactId>vision-demo</artifactId>
    <version>1.0.0</version>
    <packaging>war</packaging>
    <name>VisionDemo</name>
</project>'''

@pytest.fixture
def sample_schema_content():
    return '''CREATE TABLE IF NOT EXISTS departments (
    id          BIGINT AUTO_INCREMENT PRIMARY KEY,
    code        VARCHAR(20)  NOT NULL UNIQUE,
    name        VARCHAR(100) NOT NULL,
    parent_id   BIGINT NULL,
    sort_order  INT DEFAULT 0,
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (parent_id) REFERENCES departments(id)
);

CREATE TABLE IF NOT EXISTS employees (
    id              BIGINT AUTO_INCREMENT PRIMARY KEY,
    employee_no     VARCHAR(20)  NOT NULL UNIQUE,
    name            VARCHAR(100) NOT NULL,
    department_id   BIGINT,
    FOREIGN KEY (department_id) REFERENCES departments(id)
);'''

@pytest.fixture
def sample_entity_content():
    return '''@Entity
@Table(name = "employees")
public class Employee {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "employee_no", nullable = false, unique = true, length = 20)
    private String employeeNo;

    @Column(nullable = false, length = 100)
    private String name;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "department_id")
    private Department department;
}'''

@pytest.fixture
def sample_controller_content():
    return '''@Controller
@RequestMapping("/employee")
public class EmployeeController {

    @Autowired
    private EmployeeService employeeService;

    @GetMapping("/page")
    public String page() {
        return "fragments/employee-list";
    }

    @GetMapping("/api/search")
    @ResponseBody
    public Map<String, Object> search(
            @RequestParam(required = false) String name,
            @RequestParam(required = false) Long deptId,
            @RequestParam(defaultValue = "0") int page,
            @RequestParam(defaultValue = "10") int size) {
        Page<Employee> result = employeeService.search(name, deptId, page, size);
        Map<String, Object> map = new HashMap<>();
        map.put("content", result.getContent());
        map.put("totalPages", result.getTotalPages());
        return map;
    }

    @PostMapping("/api/save")
    @ResponseBody
    public ResponseEntity<Employee> save(@RequestBody Employee employee) {
        Employee saved = employeeService.save(employee);
        return ResponseEntity.ok(saved);
    }

    @DeleteMapping("/api/{id}")
    @ResponseBody
    public ResponseEntity<Void> delete(@PathVariable Long id) {
        employeeService.delete(id);
        return ResponseEntity.ok().build();
    }
}'''

@pytest.fixture
def sample_jsp_content():
    return '''<%@ page contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<div class="search-bar">
    <label>氏名: <input type="text" id="emp-search-name" size="16"></label>
    <label>部署:
        <select id="emp-search-dept">
            <option value="">すべて</option>
        </select>
    </label>
    <button class="btn btn-primary btn-sm" id="emp-search-btn">検索</button>
    <button class="btn btn-sm" id="emp-reset-btn">リセット</button>
</div>
<div id="emp-table-container"></div>
<div id="emp-pagination-container"></div>
<script>
$(function() {
    function doSearch(page) {
        getJSON('/employee/api/search', {
            name: $('#emp-search-name').val(),
            deptId: $('#emp-search-dept').val(),
            page: page || 0
        }, function(res) {
            TableUtils.renderTable($('#emp-table-container'), {
                columns: [
                    { key: 'employeeNo', label: '社員番号', width: '90px' },
                    { key: 'name', label: '氏名', width: '120px' },
                    { key: 'position', label: '役職', width: '80px' }
                ],
                data: res.content
            });
        });
    }
});
</script>'''
```

- [ ] **Step 4: 依存パッケージをインストール**

```bash
cd /Users/mz/Documents/01_work/EasySpec && pip install -r requirements.txt
```

- [ ] **Step 5: テストが空で通ることを確認**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/ -v
```

Expected: `no tests ran`（まだテストファイルがないため）

- [ ] **Step 6: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git init && git add -A && git commit -m "chore: initialize project structure"
```

---

### Task 2: データモデル定義

**Files:**
- Create: `src/models.py`

- [ ] **Step 1: モデルのテストを書くほど単純なので、直接実装**

```python
# src/models.py
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class SystemInfo:
    name: str
    version: str
    packaging: str = "jar"


@dataclass
class ColumnInfo:
    name: str
    logical_name: str = ""
    sql_type: str = ""
    type_name: str = ""
    length: Optional[int] = None
    nullable: bool = True
    key_type: str = ""  # PK, FK, UNIQUE, ""
    default_value: str = ""
    description: str = ""


@dataclass
class ForeignKeyInfo:
    column: str
    ref_table: str
    ref_column: str
    on_delete: str = ""


@dataclass
class TableInfo:
    name: str
    logical_name: str = ""
    columns: list = field(default_factory=list)
    foreign_keys: list = field(default_factory=list)
    description: str = ""


@dataclass
class TableColumnDef:
    key: str
    label: str
    width: str = ""
    data_type: str = "text"


@dataclass
class ScreenField:
    name: str
    field_type: str = "text"
    length: int = 0
    required: bool = False
    description: str = ""


@dataclass
class ScreenButton:
    name: str
    action: str = ""
    description: str = ""


@dataclass
class ValidationRule:
    field: str
    rule: str
    message: str = ""


@dataclass
class ScreenInfo:
    id: str
    name: str
    url: str
    jsp_file: str = ""
    controller: str = ""
    layout_description: str = ""
    fields: list = field(default_factory=list)
    table_columns: list = field(default_factory=list)
    buttons: list = field(default_factory=list)
    validations: list = field(default_factory=list)
    is_mock: bool = False


@dataclass
class ApiParam:
    name: str
    location: str  # path, query, body
    required: bool = False
    param_type: str = "String"
    description: str = ""


@dataclass
class ApiEndpoint:
    method: str
    url: str
    function_name: str = ""
    request_params: list = field(default_factory=list)
    response_type: str = ""
    controller: str = ""
    description: str = ""
```

- [ ] **Step 2: モデルのインポート確認テスト**

```python
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
```

- [ ] **Step 3: テスト実行**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_models.py -v
```

Expected: 4 tests PASS

- [ ] **Step 4: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/models.py tests/test_models.py && git commit -m "feat: add data models"
```

---

### Task 3: pom.xml パーサー

**Files:**
- Create: `src/parsers/pom_parser.py`
- Create: `tests/test_pom_parser.py`

- [ ] **Step 1: 失敗するテストを書く**

```python
# tests/test_pom_parser.py
import tempfile
import os
from src.parsers.pom_parser import parse_pom


def test_parse_pom_returns_system_info(sample_pom_content):
    with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False, encoding='utf-8') as f:
        f.write(sample_pom_content)
        tmp_path = f.name

    try:
        info = parse_pom(tmp_path)
        assert info.name == "VisionDemo"
        assert info.version == "1.0.0"
        assert info.packaging == "war"
    finally:
        os.unlink(tmp_path)


def test_parse_pom_no_name_fallback_to_artifact_id(sample_pom_content):
    content_no_name = sample_pom_content.replace('<name>VisionDemo</name>', '')
    with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False, encoding='utf-8') as f:
        f.write(content_no_name)
        tmp_path = f.name

    try:
        info = parse_pom(tmp_path)
        assert info.name == "vision-demo"
    finally:
        os.unlink(tmp_path)
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_pom_parser.py -v
```

Expected: FAIL — `ModuleNotFoundError: No module named 'src.parsers.pom_parser'`

- [ ] **Step 3: 実装**

```python
# src/parsers/pom_parser.py
import xml.etree.ElementTree as ET
from src.models import SystemInfo

NS = {'m': 'http://maven.apache.org/POM/4.0.0'}


def parse_pom(filepath: str) -> SystemInfo:
    tree = ET.parse(filepath)
    root = tree.getroot()

    name = _find_text(root, 'm:name', '')
    artifact_id = _find_text(root, 'm:artifactId', 'unknown')
    version = _find_text(root, 'm:version', '0.0.0')
    packaging = _find_text(root, 'm:packaging', 'jar')

    return SystemInfo(
        name=name or artifact_id,
        version=version,
        packaging=packaging
    )


def _find_text(root, tag, default):
    el = root.find(tag, NS)
    return el.text.strip() if el is not None and el.text else default
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_pom_parser.py -v
```

Expected: 2 tests PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/parsers/pom_parser.py tests/test_pom_parser.py && git commit -m "feat: add pom.xml parser"
```

---

### Task 4: schema.sql パーサー

**Files:**
- Create: `src/parsers/schema_parser.py`
- Create: `tests/test_schema_parser.py`

- [ ] **Step 1: 失敗するテストを書く**

```python
# tests/test_schema_parser.py
import tempfile
import os
from src.parsers.schema_parser import parse_schema


def test_parse_schema_returns_table_list(sample_schema_content):
    with tempfile.NamedTemporaryFile(mode='w', suffix='.sql', delete=False, encoding='utf-8') as f:
        f.write(sample_schema_content)
        tmp_path = f.name

    try:
        tables = parse_schema(tmp_path)
        assert len(tables) == 2

        # departments table
        dept = tables[0]
        assert dept.name == "departments"
        assert len(dept.columns) == 6
        assert dept.columns[0].name == "id"
        assert dept.columns[0].sql_type == "BIGINT"
        assert dept.columns[0].key_type == "PK"

        # code column
        assert dept.columns[1].name == "code"
        assert dept.columns[1].type_name == "VARCHAR"
        assert dept.columns[1].length == 20
        assert dept.columns[1].nullable is False
        assert "UNIQUE" in dept.columns[1].key_type

        # parent_id FK
        assert dept.columns[3].name == "parent_id"
        assert len(dept.foreign_keys) == 1
        assert dept.foreign_keys[0].column == "parent_id"
        assert dept.foreign_keys[0].ref_table == "departments"

        # employees table
        emp = tables[1]
        assert emp.name == "employees"
        assert len(emp.columns) == 4
        assert emp.columns[0].name == "id"
        assert emp.columns[0].key_type == "PK"
        assert emp.foreign_keys[0].column == "department_id"
        assert emp.foreign_keys[0].ref_table == "departments"
    finally:
        os.unlink(tmp_path)


def test_column_with_default_value():
    sql = """CREATE TABLE IF NOT EXISTS test_table (
    id BIGINT AUTO_INCREMENT PRIMARY KEY,
    status VARCHAR(20) DEFAULT '下書き'
);"""
    with tempfile.NamedTemporaryFile(mode='w', suffix='.sql', delete=False, encoding='utf-8') as f:
        f.write(sql)
        tmp_path = f.name

    try:
        tables = parse_schema(tmp_path)
        assert tables[0].columns[1].default_value == "'下書き'"
    finally:
        os.unlink(tmp_path)
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_schema_parser.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/parsers/schema_parser.py
import re
from src.models import TableInfo, ColumnInfo, ForeignKeyInfo


def parse_schema(filepath: str) -> list:
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    tables = []
    # Split by CREATE TABLE
    table_blocks = re.split(r'CREATE TABLE IF NOT EXISTS\s+', content, flags=re.IGNORECASE)
    for block in table_blocks[1:]:  # skip content before first CREATE TABLE
        table = _parse_table_block(block)
        if table:
            tables.append(table)
    return tables


def _parse_table_block(block: str) -> TableInfo:
    lines = block.split('\n')
    table_name = lines[0].strip().split('(')[0].strip()

    # Extract everything between first ( and last )
    body_start = block.index('(') + 1
    # Find the matching closing paren by looking for the last ); 
    idx = block.rfind(')')
    body = block[body_start:idx]

    columns = []
    foreign_keys = []

    # Split by comma, but be careful with function args
    statements = _split_sql_statements(body)

    for stmt in statements:
        stmt = stmt.strip()
        if not stmt:
            continue

        # FOREIGN KEY constraint
        fk_match = re.match(
            r'FOREIGN\s+KEY\s*\((\w+)\)\s*REFERENCES\s*(\w+)\s*\((\w+)\)\s*(.*)',
            stmt, re.IGNORECASE
        )
        if fk_match:
            on_delete = ''
            on_delete_match = re.search(r'ON\s+DELETE\s+(\w+)', fk_match.group(4), re.IGNORECASE)
            if on_delete_match:
                on_delete = on_delete_match.group(1)
            foreign_keys.append(ForeignKeyInfo(
                column=fk_match.group(1),
                ref_table=fk_match.group(2),
                ref_column=fk_match.group(3),
                on_delete=on_delete
            ))
            continue

        # Regular column definition
        col = _parse_column_def(stmt)
        if col:
            columns.append(col)

    return TableInfo(name=table_name, columns=columns, foreign_keys=foreign_keys)


def _parse_column_def(stmt: str) -> ColumnInfo:
    # Pattern: name TYPE[(len)] [options...]
    match = re.match(r'(\w+)\s+(\w+)(\([\d,]+\))?\s*(.*)', stmt, re.IGNORECASE | re.DOTALL)
    if not match:
        return None

    name = match.group(1)
    type_name = match.group(2).upper()
    type_len_str = match.group(3)
    options = match.group(4)

    length = None
    if type_len_str:
        nums = type_len_str.strip('()')
        # Take first number if precision,scale
        length = int(nums.split(',')[0])

    sql_type = type_name
    if type_len_str:
        sql_type = f"{type_name}{type_len_str}"

    nullable = 'NOT NULL' not in options.upper()
    key_type = _extract_key_type(options)

    default_value = ''
    default_match = re.search(r"DEFAULT\s+('[^']*'|\S+)", options, re.IGNORECASE)
    if default_match:
        default_value = default_match.group(1)
    # Handle DEFAULT CURRENT_TIMESTAMP
    if re.search(r'DEFAULT\s+CURRENT_TIMESTAMP', options, re.IGNORECASE):
        default_value = 'CURRENT_TIMESTAMP'

    return ColumnInfo(
        name=name,
        sql_type=sql_type,
        type_name=type_name,
        length=length,
        nullable=nullable,
        key_type=key_type,
        default_value=default_value,
    )


def _extract_key_type(options: str) -> str:
    keys = []
    upper = options.upper()
    if 'AUTO_INCREMENT' in upper or 'PRIMARY KEY' in upper:
        keys.append('PK')
    if 'UNIQUE' in upper:
        keys.append('UNIQUE')
    return ', '.join(keys)


def _split_sql_statements(body: str) -> list:
    """Split CREATE TABLE body by comma, respecting nested parentheses."""
    result = []
    depth = 0
    current = []
    for ch in body:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        elif ch == ',' and depth == 0:
            result.append(''.join(current))
            current = []
            continue
        current.append(ch)
    if current:
        result.append(''.join(current))
    return result
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_schema_parser.py -v
```

Expected: 2 tests PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/parsers/schema_parser.py tests/test_schema_parser.py && git commit -m "feat: add schema.sql parser"
```

---

### Task 5: Entity.java パーサー

**Files:**
- Create: `src/parsers/entity_parser.py`
- Create: `tests/test_entity_parser.py`

- [ ] **Step 1: 失敗するテストを書く**

```python
# tests/test_entity_parser.py
import tempfile
import os
from src.parsers.entity_parser import parse_entity


def test_parse_entity_extracts_table_and_columns(sample_entity_content):
    with tempfile.NamedTemporaryFile(mode='w', suffix='.java', delete=False, encoding='utf-8') as f:
        f.write(sample_entity_content)
        tmp_path = f.name

    try:
        result = parse_entity(tmp_path)
        assert result['table_name'] == "employees"
        assert len(result['columns']) == 4

        # id column
        id_col = result['columns'][0]
        assert id_col['name'] == 'id'
        assert id_col['java_type'] == 'Long'
        assert id_col['is_id'] is True

        # employee_no column with @Column annotations
        no_col = result['columns'][1]
        assert no_col['name'] == 'employee_no'
        assert no_col['java_type'] == 'String'
        assert no_col['nullable'] is False
        assert no_col['unique'] is True
        assert no_col['length'] == 20

        # name column
        name_col = result['columns'][2]
        assert name_col['name'] == 'name'
        assert name_col['nullable'] is False
        assert name_col['length'] == 100

        # department FK
        dept_col = result['columns'][3]
        assert dept_col['name'] == 'department_id'
        assert dept_col['is_fk'] is True
        assert dept_col['fk_entity'] == 'Department'

        assert len(result['foreign_keys']) == 1
        assert result['foreign_keys'][0]['column'] == 'department_id'
        assert result['foreign_keys'][0]['ref_table'] == 'departments'
        assert result['foreign_keys'][0]['ref_column'] == 'id'
    finally:
        os.unlink(tmp_path)
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_entity_parser.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/parsers/entity_parser.py
import re
import os


def parse_entity(filepath: str) -> dict:
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    class_name = os.path.splitext(os.path.basename(filepath))[0]
    table_name = _extract_table_name(content, class_name)

    columns = []
    foreign_keys = []

    # Find all field declarations with their annotations
    field_blocks = _split_fields(content)

    for block in field_blocks:
        col_info = _parse_field_block(block, table_name)
        if col_info['is_fk']:
            foreign_keys.append({
                'column': col_info['name'],
                'ref_table': col_info.get('fk_table', ''),
                'ref_column': col_info.get('fk_column', 'id'),
                'on_delete': col_info.get('on_delete', ''),
            })
        columns.append(col_info)

    return {
        'table_name': table_name,
        'entity_class': class_name,
        'columns': columns,
        'foreign_keys': foreign_keys,
    }


def _extract_table_name(content: str, class_name: str) -> str:
    # @Table(name = "employees")
    m = re.search(r'@Table\s*\(\s*name\s*=\s*"(\w+)"', content)
    if m:
        return m.group(1)
    # Fallback: lowercase class name + 's'
    return class_name.lower() + 's'


def _split_fields(content: str) -> list:
    """Split entity content into field blocks (annotation block + field declaration)."""
    # Match: annotation lines followed by field declaration
    pattern = r'((?:@\w+[^;]*\s*)*)\s*private\s+(\w+(?:<\w+>)?)\s+(\w+)\s*;'
    matches = re.findall(pattern, content)
    # Return combined annotation + field info
    result = []
    for annotations_raw, java_type, field_name in matches:
        result.append({
            'annotations': annotations_raw,
            'java_type': java_type,
            'field_name': field_name,
        })
    return result


def _parse_field_block(field_data: dict, table_name: str) -> dict:
    annotations = field_data['annotations']
    java_type = field_data['java_type']
    field_name = field_data['field_name']

    is_id = '@Id' in annotations
    is_fk = '@ManyToOne' in annotations or '@OneToMany' in annotations

    col_info = {
        'name': field_name,
        'java_type': java_type,
        'is_id': is_id,
        'is_fk': is_fk,
        'nullable': True,
        'unique': False,
        'length': None,
        'fk_entity': None,
        'fk_table': None,
        'fk_column': None,
        'on_delete': '',
    }

    # Extract @Column properties
    col_match = re.search(r'@Column\s*\(([^)]+)\)', annotations)
    if col_match:
        props = col_match.group(1)
        name_m = re.search(r'name\s*=\s*"(\w+)"', props)
        if name_m:
            col_info['name'] = name_m.group(1)

        if 'nullable' in props:
            null_m = re.search(r'nullable\s*=\s*(true|false)', props)
            if null_m:
                col_info['nullable'] = null_m.group(1) == 'true'

        if 'unique' in props:
            uniq_m = re.search(r'unique\s*=\s*(true|false)', props)
            if uniq_m:
                col_info['unique'] = uniq_m.group(1) == 'true'

        len_m = re.search(r'length\s*=\s*(\d+)', props)
        if len_m:
            col_info['length'] = int(len_m.group(1))

    # Extract @JoinColumn for FK
    join_match = re.search(r'@JoinColumn\s*\(([^)]+)\)', annotations)
    if join_match:
        props = join_match.group(1)
        name_m = re.search(r'name\s*=\s*"(\w+)"', props)
        if name_m:
            col_info['name'] = name_m.group(1)

    # FK entity
    if is_fk and '@ManyToOne' in annotations:
        # The field type is the entity class
        col_info['fk_entity'] = java_type
        col_info['fk_table'] = java_type.lower() + 's'

    return col_info
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_entity_parser.py -v
```

Expected: 1 test PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/parsers/entity_parser.py tests/test_entity_parser.py && git commit -m "feat: add entity.java parser"
```

---

### Task 6: Controller.java パーサー

**Files:**
- Create: `src/parsers/controller_parser.py`
- Create: `tests/test_controller_parser.py`

- [ ] **Step 1: 失敗するテストを書く**

```python
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
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_controller_parser.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/parsers/controller_parser.py
import re
import os
from src.models import ApiEndpoint, ApiParam


def parse_controller(filepath: str) -> tuple:
    """
    Returns (list[ApiEndpoint], dict[str, str] url_to_jsp_map)
    """
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    class_name = os.path.splitext(os.path.basename(filepath))[0]

    # Extract @RequestMapping base path
    base_path = ""
    req_map = re.search(r'@RequestMapping\s*\(\s*"([^"]+)"', content)
    if req_map:
        base_path = req_map.group(1)

    endpoints = []
    jsp_map = {}

    # Find all HTTP method annotated methods
    # Pattern matches: @XxxMapping(...) up to the method closing brace
    methods = _extract_method_blocks(content)

    for method_block in methods:
        endpoint = _parse_method_block(method_block, base_path, class_name)
        if endpoint:
            endpoints.append(endpoint)
            # Check if it returns a JSP view name
            jsp = _extract_jsp_view(method_block)
            if jsp:
                jsp_map[endpoint.url] = jsp

    return endpoints, jsp_map


def _extract_method_blocks(content: str) -> list:
    """Extract individual method blocks with their annotations."""
    # Match @RequestMapping, @GetMapping, @PostMapping, @PutMapping, @DeleteMapping
    # followed by the method declaration and body
    pattern = r'(@(?:Get|Post|Put|Delete|Request)Mapping[^)]*\)(?:\s*@\w+[^)]*\))*\s*public\s+.*?\{[^}]*\}'
    matches = []
    for match in re.finditer(pattern, content, re.DOTALL):
        matches.append(match.group(0))
    return matches


def _parse_method_block(block: str, base_path: str, class_name: str) -> ApiEndpoint:
    # Extract HTTP method
    method = "GET"
    if '@PostMapping' in block:
        method = "POST"
    elif '@PutMapping' in block:
        method = "PUT"
    elif '@DeleteMapping' in block:
        method = "DELETE"

    # Extract URL path
    url = ""
    for mapping_type in ['GetMapping', 'PostMapping', 'PutMapping', 'DeleteMapping', 'RequestMapping']:
        url_match = re.search(rf'@{mapping_type}\s*\(\s*"([^"]+)"', block)
        if url_match:
            url = url_match.group(1)
            break

    # Combine with base path
    full_url = _join_paths(base_path, url)

    # Extract request params
    params = []
    # @RequestParam
    for param_match in re.finditer(
        r'@RequestParam\s*\(([^)]*)\)\s*(\w+(?:<\w+>)?)\s+(\w+)',
        block
    ):
        props = param_match.group(1)
        param_type = param_match.group(2)
        param_name = param_match.group(3)
        required = 'required' not in props or 'false' not in props.split('=')[-1].strip()
        params.append(ApiParam(
            name=param_name,
            location="query",
            required=required,
            param_type=param_type,
        ))

    # @PathVariable
    for path_match in re.finditer(
        r'@PathVariable\s*(?:\([^)]*\))?\s*(\w+(?:<\w+>)?)\s+(\w+)',
        block
    ):
        params.append(ApiParam(
            name=path_match.group(2),
            location="path",
            required=True,
            param_type=path_match.group(1),
        ))

    # @RequestBody
    for body_match in re.finditer(
        r'@RequestBody\s+(\w+(?:<\w+>)?)\s+(\w+)',
        block
    ):
        params.append(ApiParam(
            name=body_match.group(2),
            location="body",
            required=True,
            param_type=body_match.group(1),
        ))

    # Response type
    response_type = ""
    ret_match = re.search(r'public\s+(\w+(?:<\w+[,\s\w<>]*>)?)\s+\w+\s*\(', block)
    if ret_match:
        response_type = ret_match.group(1)

    return ApiEndpoint(
        method=method,
        url=full_url,
        controller=class_name,
        request_params=params,
        response_type=response_type,
    )


def _extract_jsp_view(block: str) -> str:
    """Extract JSP view name from return statement."""
    m = re.search(r'return\s*"([^"]+)"', block)
    if m:
        return m.group(1)
    return ""


def _join_paths(base: str, sub: str) -> str:
    if not base:
        return sub
    if not sub:
        return base
    base = base.rstrip('/')
    sub = sub.lstrip('/')
    return f"{base}/{sub}"
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_controller_parser.py -v
```

Expected: 1 test PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/parsers/controller_parser.py tests/test_controller_parser.py && git commit -m "feat: add controller.java parser"
```

---

### Task 7: JSP パーサー

**Files:**
- Create: `src/parsers/jsp_parser.py`
- Create: `tests/test_jsp_parser.py`

- [ ] **Step 1: 失敗するテストを書く**

```python
# tests/test_jsp_parser.py
import tempfile
import os
from src.parsers.jsp_parser import parse_jsp


def test_parse_jsp_extracts_fields_buttons_columns(sample_jsp_content):
    with tempfile.NamedTemporaryFile(mode='w', suffix='.jsp', delete=False, encoding='utf-8') as f:
        f.write(sample_jsp_content)
        tmp_path = f.name

    try:
        result = parse_jsp(tmp_path)

        # Fields (input elements)
        assert len(result['fields']) >= 1
        search_name = [f for f in result['fields'] if f.name == 'emp-search-name'][0]
        assert search_name.field_type == 'text'

        # Buttons
        assert len(result['buttons']) >= 2
        btn_names = [b.name for b in result['buttons']]
        assert 'emp-search-btn' in btn_names
        assert 'emp-reset-btn' in btn_names

        # Table columns (from renderTable)
        assert len(result['table_columns']) == 3
        assert result['table_columns'][0].key == 'employeeNo'
        assert result['table_columns'][0].label == '社員番号'
        assert result['table_columns'][1].key == 'name'
        assert result['table_columns'][1].label == '氏名'
        assert result['table_columns'][1].width == '120px'

        # Layout (has search bar + table)
        assert 'search-bar' in result['layout_description'] or result['has_search'] is True
        assert result['has_table'] is True
    finally:
        os.unlink(tmp_path)
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_jsp_parser.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/parsers/jsp_parser.py
import re
from src.models import ScreenField, ScreenButton, TableColumnDef


def parse_jsp(filepath: str) -> dict:
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    fields = _extract_input_fields(content)
    buttons = _extract_buttons(content)
    table_columns = _extract_table_columns(content)
    has_search = bool(re.search(r'search-bar|search-btn', content))
    has_table = bool(re.search(r'table-container|renderTable', content))
    layout_parts = _describe_layout(content)

    return {
        'fields': fields,
        'buttons': buttons,
        'table_columns': table_columns,
        'has_search': has_search,
        'has_table': has_table,
        'layout_description': ' / '.join(layout_parts) if layout_parts else '',
    }


def _extract_input_fields(content: str) -> list:
    fields = []
    # Match <input> elements
    for m in re.finditer(
        r'<input\s+[^>]*?(?:id|name)\s*=\s*"(\w+(?:-\w+)*)"[^>]*?(?:type\s*=\s*"(\w+)")?[^>]*>',
        content
    ):
        input_id = m.group(1)
        input_type = m.group(2) or 'text'
        fields.append(ScreenField(
            name=input_id,
            field_type=input_type,
        ))
    # Match <select> elements
    for m in re.finditer(r'<select\s+[^>]*?id\s*=\s*"(\w+(?:-\w+)*)"', content):
        fields.append(ScreenField(
            name=m.group(1),
            field_type='select',
        ))
    # Match <textarea> elements
    for m in re.finditer(r'<textarea\s+[^>]*?id\s*=\s*"(\w+(?:-\w+)*)"', content):
        fields.append(ScreenField(
            name=m.group(1),
            field_type='textarea',
        ))
    return fields


def _extract_buttons(content: str) -> list:
    buttons = []
    for m in re.finditer(
        r'<button\s+[^>]*?id\s*=\s*"(\w+(?:-\w+)*)"[^>]*>(.*?)</button>',
        content, re.DOTALL
    ):
        label = re.sub(r'<[^>]+>', '', m.group(2)).strip()
        buttons.append(ScreenButton(
            name=m.group(1),
            action=label,
        ))
    return buttons


def _extract_table_columns(content: str) -> list:
    """Extract table column definitions from JS renderTable calls."""
    columns = []
    # Match { key: 'xxx', label: 'xxx', width: 'xxx' } patterns
    pattern = r"\{\s*key\s*:\s*'(\w+)'\s*,\s*label\s*:\s*'([^']+)'\s*(?:,\s*width\s*:\s*'([^']+)')?"
    for m in re.finditer(pattern, content):
        columns.append(TableColumnDef(
            key=m.group(1),
            label=m.group(2),
            width=m.group(3) or '',
        ))
    return columns


def _describe_layout(content: str) -> list:
    parts = []
    if 'search-bar' in content or 'search-btn' in content:
        parts.append('検索エリア')
    if 'table-container' in content or 'renderTable' in content:
        parts.append('一覧テーブル')
    if 'pagination' in content:
        parts.append('ページネーション')
    if 'modal' in content.lower() or 'showModal' in content:
        parts.append('モーダル編集')
    if 'tab-item' in content:
        parts.append('タブ切替')
    if 'tree-toggle' in content or 'tree-item' in content:
        parts.append('ツリー表示')
    if 'wizard' in content.lower() or 'wiz-step' in content:
        parts.append('ウィザード形式')
    return parts
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_jsp_parser.py -v
```

Expected: 1 test PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/parsers/jsp_parser.py tests/test_jsp_parser.py && git commit -m "feat: add jsp parser"
```

---

### Task 8: project_reader オーケストレータ

**Files:**
- Create: `src/project_reader.py`
- Create: `tests/test_project_reader.py`

- [ ] **Step 1: 失敗するテストを書く**

```python
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
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_project_reader.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/project_reader.py
import os
import glob
from src.models import ScreenInfo
from src.parsers.pom_parser import parse_pom
from src.parsers.schema_parser import parse_schema
from src.parsers.entity_parser import parse_entity
from src.parsers.controller_parser import parse_controller
from src.parsers.jsp_parser import parse_jsp


class ProjectReader:
    def __init__(self, project_path: str):
        self.project_path = project_path
        self.java_src = os.path.join(project_path, 'src', 'main', 'java')
        self.webapp = os.path.join(project_path, 'src', 'main', 'webapp', 'WEB-INF', 'jsp')
        self.resources = os.path.join(project_path, 'src', 'main', 'resources')

    def read_system_info(self):
        pom_path = os.path.join(self.project_path, 'pom.xml')
        return parse_pom(pom_path)

    def read_tables(self) -> list:
        schema_path = os.path.join(self.resources, 'schema.sql')
        tables = parse_schema(schema_path)

        # Supplement with entity annotation info
        entity_dir = self._find_package_dir('entity')
        entity_infos = {}
        if entity_dir:
            for java_file in glob.glob(os.path.join(entity_dir, '*.java')):
                entity_info = parse_entity(java_file)
                entity_infos[entity_info['table_name']] = entity_info

        # Merge entity annotations into schema tables
        for table in tables:
            if table.name in entity_infos:
                ei = entity_infos[table.name]
                self._merge_entity_info(table, ei)

        # Add logical names
        self._assign_logical_names(tables)
        return tables

    def _merge_entity_info(self, table, entity_info):
        """Merge @Column annotations into schema-derived table info."""
        for col in table.columns:
            for ei_col in entity_info['columns']:
                if ei_col['name'] == col.name:
                    if ei_col.get('length') and not col.length:
                        col.length = ei_col['length']
                    if 'nullable' in ei_col and ei_col['nullable'] is False:
                        col.nullable = False
        # Add entity-derived FK info if not already present
        existing_fk_cols = {fk.column for fk in table.foreign_keys}
        for ei_fk in entity_info['foreign_keys']:
            if ei_fk['column'] not in existing_fk_cols:
                from src.models import ForeignKeyInfo
                table.foreign_keys.append(ForeignKeyInfo(
                    column=ei_fk['column'],
                    ref_table=ei_fk['ref_table'],
                    ref_column=ei_fk['ref_column'],
                    on_delete=ei_fk.get('on_delete', ''),
                ))

    def _assign_logical_names(self, tables):
        logical_map = {
            'departments': '部署マスタ',
            'employees': '社員情報',
            'expenses': '経費精算',
            'expense_details': '経費明細',
            'orders': '発注',
            'order_items': '発注明細',
        }
        for table in tables:
            table.logical_name = logical_map.get(table.name, table.name)

        column_logical_map = {
            'departments': {
                'id': ('ID', '主キー'),
                'code': ('部署コード', '部署を一意に識別するコード'),
                'name': ('部署名', '部署の正式名称'),
                'parent_id': ('上位部署ID', '自己参照による上位部署'),
                'sort_order': ('表示順', '同一階層内の表示順序'),
                'created_at': ('作成日時', 'レコード作成日時'),
            },
            'employees': {
                'id': ('ID', '主キー'),
                'employee_no': ('社員番号', '社員を一意に識別する番号'),
                'name': ('氏名', '社員の氏名'),
                'name_kana': ('フリガナ', '氏名のフリガナ'),
                'department_id': ('部署ID', '所属部署の外部キー'),
                'position': ('役職', '課長/部長/主任など'),
                'hire_date': ('入社日', '入社年月日'),
                'email': ('メール', '会社メールアドレス'),
                'phone': ('電話', '内線または外線番号'),
                'created_at': ('作成日時', 'レコード作成日時'),
                'updated_at': ('更新日時', 'レコード更新日時'),
            },
            'expenses': {
                'id': ('ID', '主キー'),
                'expense_no': ('経費番号', '精算申請番号'),
                'employee_id': ('社員ID', '申請者の社員外部キー'),
                'total_amount': ('合計金額', '経費合計額'),
                'status': ('ステータス', '下書き/申請中/承認済/差戻し'),
                'apply_date': ('申請日', '経費発生日または申請日'),
                'description': ('摘要', '申請内容の概要'),
                'created_at': ('作成日時', 'レコード作成日時'),
                'updated_at': ('更新日時', 'レコード更新日時'),
            },
            'expense_details': {
                'id': ('ID', '主キー'),
                'expense_id': ('経費ID', '経費精算の外部キー'),
                'line_no': ('行番号', '明細行の連番'),
                'account_item': ('科目', '交通費/宿泊費/消耗品費 等'),
                'amount': ('金額', '明細行の金額'),
                'expense_date': ('発生日', '経費発生日'),
                'description': ('摘要', '明細の説明'),
            },
            'orders': {
                'id': ('ID', '主キー'),
                'order_no': ('発注番号', '発注を一意に識別する番号'),
                'supplier': ('仕入先', '発注先企業名'),
                'total_amount': ('合計金額', '発注合計額'),
                'status': ('ステータス', '新規/発注済/納品済/キャンセル'),
                'order_date': ('発注日', '発注日'),
                'delivery_date': ('納品日', '納品予定日または納品日'),
                'created_at': ('作成日時', 'レコード作成日時'),
                'updated_at': ('更新日時', 'レコード更新日時'),
            },
            'order_items': {
                'id': ('ID', '主キー'),
                'order_id': ('発注ID', '発注の外部キー'),
                'line_no': ('行番号', '明細行の連番'),
                'item_name': ('品名', '発注品名'),
                'quantity': ('数量', '発注数量'),
                'unit_price': ('単価', '単価'),
            },
        }

        for table in tables:
            col_map = column_logical_map.get(table.name, {})
            for col in table.columns:
                if col.name in col_map:
                    col.logical_name, col.description = col_map[col.name]

    def read_api_endpoints(self) -> tuple:
        """Returns (list[ApiEndpoint], dict[str,str] jsp_map)"""
        controller_dir = self._find_package_dir('controller')
        all_endpoints = []
        all_jsp_map = {}
        if controller_dir:
            for java_file in sorted(glob.glob(os.path.join(controller_dir, '*.java'))):
                endpoints, jsp_map = parse_controller(java_file)
                all_endpoints.extend(endpoints)
                all_jsp_map.update(jsp_map)
        return all_endpoints, all_jsp_map

    def read_screens(self) -> list:
        """Read all screens combining controller + JSP info."""
        endpoints, jsp_map = self.read_api_endpoints()

        screen_defs = [
            ('SCR-001', 'メインメニュー', '/', 'index.jsp', 'IndexController',
             'サイドバーツリーメニュー + ワークスペース + タスクバー'),
            ('SCR-002', '組織管理', '/org/page', 'fragments/org-tree.jsp', 'OrgController',
             '部署ツリー表示、ドラッグ＆ドロップ移動、インライン編集、右クリックメニュー'),
            ('SCR-003', '社員管理', '/employee/page', 'fragments/employee-list.jsp', 'EmployeeController',
             '社員一覧検索・表示、モーダル編集、削除'),
            ('SCR-004', '経費精算一覧', '/expense/page?view=list', 'fragments/expense-list.jsp', 'ExpenseController',
             '経費精算一覧、ステータス絞込、モーダル詳細表示'),
            ('SCR-005', '経費精算登録', '/expense/page?view=create', 'fragments/expense-create.jsp', 'ExpenseController',
             '3ステップウィザード形式の経費精算登録（基本情報→明細入力→確認→申請）'),
            ('SCR-006', '発注管理', '/order/page', 'fragments/order-list.jsp', 'OrderController',
             '発注一覧検索、詳細モーダル、キャンセル・複製操作、右クリックメニュー'),
            ('SCR-007', 'レポート', '/report/page', 'fragments/report.jsp', 'ReportController',
             '部署別・社員別経費/発注集計表、モックデータ生成'),
            ('SCR-008', 'システム辞書', '/system/page?view=dict', 'fragments/system-dict.jsp', 'SystemController',
             '辞書項目ツリー管理、サンプル実装'),
            ('SCR-009', 'システムパラメータ', '/system/page?view=param', 'fragments/system-param.jsp', 'SystemController',
             'システムパラメータ設定（メール通知/自動バックアップ/セッションタイムアウト 等）'),
        ]

        screens = []
        for sid, name, url, jsp, ctrl, layout in screen_defs:
            screen = ScreenInfo(
                id=sid, name=name, url=url,
                jsp_file=jsp, controller=ctrl,
                layout_description=layout,
            )
            # Parse JSP for fields/buttons/columns
            jsp_path = os.path.join(self.webapp, jsp)
            if os.path.exists(jsp_path):
                parsed = parse_jsp(jsp_path)
                screen.fields = parsed['fields']
                screen.buttons = parsed['buttons']
                screen.table_columns = parsed['table_columns']
                if parsed['layout_description']:
                    screen.layout_description += ' | ' + parsed['layout_description']
            # Mark mock implementations
            if ctrl == 'ReportController' or ctrl == 'SystemController':
                screen.is_mock = True
            screens.append(screen)

        return screens

    def read_entity_infos(self) -> dict:
        """Returns {table_name: entity_info} for all entities."""
        entity_dir = self._find_package_dir('entity')
        result = {}
        if entity_dir:
            for java_file in glob.glob(os.path.join(entity_dir, '*.java')):
                info = parse_entity(java_file)
                result[info['table_name']] = info
        return result

    def _find_package_dir(self, subpackage: str) -> str:
        """Find the Java package directory under src/main/java."""
        base = self.java_src
        if not os.path.isdir(base):
            return None
        for root, dirs, files in os.walk(base):
            if os.path.basename(root) == subpackage:
                return root
        return None
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_project_reader.py -v
```

Expected: 1 test PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/project_reader.py tests/test_project_reader.py && git commit -m "feat: add project reader orchestrator"
```

---

### Task 9: Excel スタイル定義

**Files:**
- Create: `src/generators/styles.py`
- Create: `tests/test_styles.py`

- [ ] **Step 1: テストを書く**

```python
# tests/test_styles.py
from openpyxl import Workbook
from src.generators.styles import (
    create_styles, apply_header_style, apply_body_style,
    apply_title_style, apply_border_thin
)


def test_create_styles_returns_dict():
    styles = create_styles()
    assert 'header_font' in styles
    assert 'body_font' in styles
    assert 'title_font' in styles
    assert 'header_fill' in styles
    assert 'thin_border' in styles
    assert 'thick_border' in styles
    assert 'header_alignment' in styles
    assert 'body_alignment' in styles
    assert 'center_alignment' in styles


def test_apply_header_style():
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Test'
    styles = create_styles()
    apply_header_style(ws['A1'], styles)
    assert ws['A1'].font.bold is True
    assert ws['A1'].fill.start_color.rgb == '00DCE6F1'


def test_apply_body_style():
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Test'
    styles = create_styles()
    apply_body_style(ws['A1'], styles)
    assert ws['A1'].font.size == 10


def test_apply_title_style():
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Title'
    styles = create_styles()
    apply_title_style(ws['A1'], styles)
    assert ws['A1'].font.size == 14
    assert ws['A1'].font.bold is True
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_styles.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/generators/styles.py
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)


def create_styles() -> dict:
    return {
        'title_font': Font(name='Yu Gothic', size=14, bold=True),
        'header_font': Font(name='Yu Gothic', size=10, bold=True),
        'body_font': Font(name='Yu Gothic', size=10),
        'small_font': Font(name='Yu Gothic', size=9),
        'header_fill': PatternFill(start_color='DCE6F1', end_color='DCE6F1', fill_type='solid'),
        'thin_border': Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
        ),
        'thick_outside': Border(
            left=Side(style='medium'),
            right=Side(style='medium'),
            top=Side(style='medium'),
            bottom=Side(style='medium'),
        ),
        'header_alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'body_alignment': Alignment(vertical='center', wrap_text=True),
        'center_alignment': Alignment(horizontal='center', vertical='center'),
    }


def apply_header_style(cell, styles):
    cell.font = styles['header_font']
    cell.fill = styles['header_fill']
    cell.border = styles['thin_border']
    cell.alignment = styles['header_alignment']


def apply_body_style(cell, styles):
    cell.font = styles['body_font']
    cell.border = styles['thin_border']
    cell.alignment = styles['body_alignment']


def apply_title_style(cell, styles):
    cell.font = styles['title_font']
    cell.alignment = Alignment(horizontal='left', vertical='center')


def apply_thick_outside(ws, min_row, max_row, min_col, max_col, styles):
    """Apply thick border to the outside of a range."""
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            border = Border(
                left=Side(style='medium') if col == min_col else Side(style='thin'),
                right=Side(style='medium') if col == max_col else Side(style='thin'),
                top=Side(style='medium') if row == min_row else Side(style='thin'),
                bottom=Side(style='medium') if row == max_row else Side(style='thin'),
            )
            cell.border = border


def write_header_row(ws, row, headers, styles, start_col=1):
    """Write a header row with header styling."""
    for i, header in enumerate(headers):
        cell = ws.cell(row=row, column=start_col + i, value=header)
        apply_header_style(cell, styles)


def write_data_row(ws, row, data, styles, start_col=1):
    """Write a data row with body styling."""
    for i, value in enumerate(data):
        cell = ws.cell(row=row, column=start_col + i, value=value)
        apply_body_style(cell, styles)
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_styles.py -v
```

Expected: 4 tests PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/generators/styles.py tests/test_styles.py && git commit -m "feat: add excel style definitions"
```

---

### Task 10: 表紙シート生成

**Files:**
- Create: `src/generators/cover.py`
- Create: `tests/test_cover.py`

- [ ] **Step 1: テストを書く**

```python
# tests/test_cover.py
from openpyxl import Workbook
from src.models import SystemInfo
from src.generators.cover import generate_cover
from src.generators.styles import create_styles


def test_generate_cover_creates_sheet():
    wb = Workbook()
    ws = wb.active
    ws.title = "表紙"

    sys_info = SystemInfo(name="TestApp", version="1.0.0", packaging="war")
    styles = create_styles()
    generate_cover(ws, sys_info, styles)

    assert ws['A1'].value == "システム詳細設計書"
    assert "TestApp" in str(ws['A3'].value)
    assert "1.0.0" in str(ws['A4'].value)

    # Document management table headers
    assert ws['A7'].value == "版数"
    assert ws['B7'].value == "更新日"
    assert ws['C7'].value == "更新内容"
    assert ws['D7'].value == "作成者"

    # First record
    assert ws['A8'].value == "1.0"
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_cover.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/generators/cover.py
from datetime import date
from src.generators.styles import apply_title_style, write_header_row, write_data_row, apply_body_style


def generate_cover(ws, sys_info, styles):
    # Column widths
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 18

    # Title
    cell = ws['A1']
    cell.value = "システム詳細設計書"
    apply_title_style(cell, styles)
    ws.merge_cells('A1:D1')

    # System info
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

    # Blank row
    row += 1

    # Document management table
    headers = ["版数", "更新日", "更新内容", "作成者"]
    write_header_row(ws, row, headers, styles)
    row += 1

    records = [
        ("1.0", date.today().strftime('%Y-%m-%d'), "初版作成", "自動生成"),
    ]
    for rec in records:
        write_data_row(ws, row, rec, styles)
        row += 1
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_cover.py -v
```

Expected: 1 test PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/generators/cover.py tests/test_cover.py && git commit -m "feat: add cover sheet generator"
```

---

### Task 11: 画面一覧シート生成

**Files:**
- Create: `src/generators/screen_list.py`
- Create: `tests/test_screen_list.py`

- [ ] **Step 1: テストを書く**

```python
# tests/test_screen_list.py
from openpyxl import Workbook
from src.models import ScreenInfo
from src.generators.screen_list import generate_screen_list
from src.generators.styles import create_styles


def test_generate_screen_list():
    wb = Workbook()
    ws = wb.active

    screens = [
        ScreenInfo(id="SCR-001", name="メインメニュー", url="/", jsp_file="index.jsp", controller="IndexController"),
        ScreenInfo(id="SCR-002", name="組織管理", url="/org/page", jsp_file="fragments/org-tree.jsp", controller="OrgController"),
    ]
    styles = create_styles()
    generate_screen_list(ws, screens, styles)

    assert ws['A1'].value == "画面一覧"
    headers = [ws.cell(row=2, column=c).value for c in range(1, 6)]
    assert headers == ["画面ID", "画面名", "URL", "JSPファイル", "機能概要"]
    assert ws.cell(row=3, column=1).value == "SCR-001"
    assert ws.cell(row=3, column=2).value == "メインメニュー"
    assert ws.cell(row=4, column=1).value == "SCR-002"
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_screen_list.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/generators/screen_list.py
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
                screen.layout_description if screen.layout_description else _describe(screen)]
        write_data_row(ws, row, data, styles)
        row += 1


def _describe(screen):
    if "一覧" in screen.name:
        return "一覧表示・検索機能"
    if "登録" in screen.name or "作成" in screen.name:
        return "データ登録機能"
    if "メニュー" in screen.name:
        return "メインメニュー画面"
    return ""
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_screen_list.py -v
```

Expected: 1 test PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/generators/screen_list.py tests/test_screen_list.py && git commit -m "feat: add screen list sheet generator"
```

---

### Task 12: テーブル定義シート生成

**Files:**
- Create: `src/generators/table_def.py`
- Create: `tests/test_table_def.py`

- [ ] **Step 1: テストを書く**

```python
# tests/test_table_def.py
from openpyxl import Workbook
from src.models import TableInfo, ColumnInfo, ForeignKeyInfo
from src.generators.table_def import generate_table_def
from src.generators.styles import create_styles


def test_generate_table_def():
    wb = Workbook()
    ws = wb.active

    tables = [
        TableInfo(
            name="employees",
            logical_name="社員情報",
            columns=[
                ColumnInfo(name="id", logical_name="ID", sql_type="BIGINT", type_name="BIGINT",
                          nullable=False, key_type="PK", description="主キー"),
                ColumnInfo(name="employee_no", logical_name="社員番号", sql_type="VARCHAR(20)", type_name="VARCHAR",
                          length=20, nullable=False, key_type="UNIQUE", description="一意の社員番号"),
                ColumnInfo(name="name", logical_name="氏名", sql_type="VARCHAR(100)", type_name="VARCHAR",
                          length=100, nullable=False),
                ColumnInfo(name="department_id", logical_name="部署ID", sql_type="BIGINT", type_name="BIGINT",
                          nullable=True, description="所属部署"),
            ],
            foreign_keys=[
                ForeignKeyInfo(column="department_id", ref_table="departments", ref_column="id"),
            ]
        ),
    ]
    styles = create_styles()
    generate_table_def(ws, tables, styles)

    assert "社員情報" in str(ws['A1'].value)
    assert ws.cell(row=3, column=1).value == "No"
    headers = [ws.cell(row=3, column=c).value for c in range(1, 9)]
    assert "論理名" in headers
    assert "物理名" in headers

    # First data row
    assert ws.cell(row=4, column=1).value == 1
    assert ws.cell(row=4, column=2).value == "ID"
    assert ws.cell(row=4, column=3).value == "id"
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_table_def.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/generators/table_def.py
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
            row += 1  # blank row between tables

        # Title
        title = f"テーブル: {table.name} ({table.logical_name})"
        cell = ws.cell(row=row, column=1, value=title)
        apply_title_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        row += 1

        # Description
        if table.description:
            ws.cell(row=row, column=1, value=table.description)
            row += 1

        # Column headers
        write_header_row(ws, row, headers, styles)
        row += 1

        # Column data
        for ci, col in enumerate(table.columns):
            null_label = "" if col.nullable else "NOT NULL"
            data = [
                ci + 1,
                col.logical_name or col.name,
                col.name,
                col.sql_type,
                col.length if col.length else "",
                null_label,
                col.key_type,
                col.description or "",
            ]
            write_data_row(ws, row, data, styles)
            row += 1

        # FK section
        if table.foreign_keys:
            fk_header = "外部キー関連:"
            cell = ws.cell(row=row, column=1, value=fk_header)
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
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_table_def.py -v
```

Expected: 1 test PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/generators/table_def.py tests/test_table_def.py && git commit -m "feat: add table definition sheet generator"
```

---

### Task 13: ER図シート生成

**Files:**
- Create: `src/generators/er_diagram.py`
- Create: `tests/test_er_diagram.py`

- [ ] **Step 1: テストを書く**

```python
# tests/test_er_diagram.py
from openpyxl import Workbook
from src.models import TableInfo
from src.generators.er_diagram import generate_er_diagram
from src.generators.styles import create_styles


def test_generate_er_diagram():
    wb = Workbook()
    ws = wb.active

    tables = [
        TableInfo(name="departments", logical_name="部署マスタ"),
        TableInfo(name="employees", logical_name="社員情報"),
    ]
    relationships = [
        ("departments", "employees", "1", "N", "department_id"),
        ("departments", "departments", "1", "N", "parent_id"),
    ]
    styles = create_styles()
    generate_er_diagram(ws, tables, relationships, styles)

    assert ws['A1'].value == "ER図（エンティティ関連図）"
    # Check that table names appear
    content = str(ws['A5'].value or '') + str(ws['B5'].value or '')
    assert len([c for row in ws.iter_rows() for c in row if c.value and 'departments' in str(c.value)]) >= 1
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_er_diagram.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/generators/er_diagram.py
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

    # Legend
    ws.cell(row=row, column=1, value="凡例:").font = styles['header_font']
    row += 1
    ws.cell(row=row, column=1, value="■ = エンティティ（テーブル）").font = styles['body_font']
    row += 1
    ws.cell(row=row, column=1, value="→ = リレーション（外部キー）").font = styles['body_font']
    row += 2

    # Table boxes
    table_positions = {}  # table_name -> (start_row, end_row)
    for table in tables:
        start_row = row
        cell = ws.cell(row=row, column=1, value=f"■ {table.name}")
        apply_header_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        row += 1
        cell = ws.cell(row=row, column=1, value=table.logical_name)
        apply_body_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        row += 1

        # List a few key columns
        for col in table.columns[:5]:
            key_mark = "🔑" if "PK" in col.key_type else "  "
            cell = ws.cell(row=row, column=1, value=f"{key_mark} {col.name}")
            apply_body_style(cell, styles)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            row += 1

        end_row = row - 1
        table_positions[table.name] = (start_row, end_row)
        row += 1  # gap

    # Relationships section
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
        data = [parent, child, card_p, card_c, fk_col,
                f"{child}.{fk_col} → {parent}.id"]
        for i, val in enumerate(data):
            cell = ws.cell(row=row, column=i+1, value=val)
            apply_body_style(cell, styles)
        row += 1


def build_relationships(tables) -> list:
    """Build relationship list from table foreign keys."""
    relationships = []
    for table in tables:
        for fk in table.foreign_keys:
            relationships.append((
                fk.ref_table,
                table.name,
                "1",
                "N",
                fk.column,
            ))
    return relationships
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_er_diagram.py -v
```

Expected: 1 test PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/generators/er_diagram.py tests/test_er_diagram.py && git commit -m "feat: add er diagram sheet generator"
```

---

### Task 14: 画面仕様シート生成

**Files:**
- Create: `src/generators/screen_spec.py`
- Create: `tests/test_screen_spec.py`

- [ ] **Step 1: テストを書く**

```python
# tests/test_screen_spec.py
from openpyxl import Workbook
from src.models import ScreenInfo, ScreenField, ScreenButton, TableColumnDef, ValidationRule
from src.generators.screen_spec import generate_screen_specs
from src.generators.styles import create_styles


def test_generate_screen_specs():
    wb = Workbook()
    # Remove default sheet, we'll create our own
    screens = [
        ScreenInfo(
            id="SCR-003", name="社員管理", url="/employee/page",
            jsp_file="fragments/employee-list.jsp", controller="EmployeeController",
            layout_description="検索エリア + 一覧テーブル + ページネーション + モーダル編集",
            fields=[ScreenField(name="emp-search-name", field_type="text", description="氏名検索")],
            buttons=[ScreenButton(name="emp-search-btn", action="検索")],
            table_columns=[
                TableColumnDef(key="employeeNo", label="社員番号", width="90px"),
                TableColumnDef(key="name", label="氏名", width="120px"),
            ],
            validations=[ValidationRule(field="emp-search-name", rule="任意入力")],
        ),
    ]
    styles = create_styles()
    generate_screen_specs(wb, screens, styles)

    # Get the sheet we created
    ws = wb["SCR-003_社員管理"]
    assert ws is not None
    assert ws['A1'].value == "画面仕様書"
    assert "SCR-003" in str(ws['A3'].value)
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_screen_spec.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/generators/screen_spec.py
from openpyxl import Workbook
from src.generators.styles import (
    apply_title_style, apply_header_style, apply_body_style,
    write_header_row, write_data_row
)


def generate_screen_specs(wb, screens, styles):
    for screen in screens:
        sheet_name = f"{screen.id}_{screen.name}"[:31]  # Excel sheet name limit
        ws = wb.create_sheet(title=sheet_name)

        # Column widths
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 16
        ws.column_dimensions['F'].width = 40

        row = 1

        # Title
        cell = ws.cell(row=row, column=1, value=f"画面仕様書")
        apply_title_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1

        # Basic info
        info_items = [
            ("画面ID", screen.id),
            ("画面名", screen.name),
            ("URL", screen.url),
            ("JSPファイル", screen.jsp_file),
            ("Controller", screen.controller),
        ]
        for label, value in info_items:
            cell = ws.cell(row=row, column=1, value=label)
            apply_header_style(cell, styles)
            cell = ws.cell(row=row, column=2, value=value)
            apply_body_style(cell, styles)
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
            row += 1

        # Mock note
        if screen.is_mock:
            cell = ws.cell(row=row, column=1, value="※ サンプル実装 / 本番未実装")
            cell.font = styles['body_font']
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            row += 1

        row += 1

        # Layout
        cell = ws.cell(row=row, column=1, value="画面レイアウト")
        apply_header_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1
        cell = ws.cell(row=row, column=1, value=screen.layout_description)
        apply_body_style(cell, styles)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 2

        # Table columns (for list screens)
        if screen.table_columns:
            _write_section(ws, row, "表示項目一覧", styles)
            row += 1
            write_header_row(ws, row, ["No", "項目キー", "表示ラベル", "幅", "データ型", ""], styles)
            row += 1
            for i, col in enumerate(screen.table_columns):
                write_data_row(ws, row, [i+1, col.key, col.label, col.width, col.data_type, ""], styles)
                row += 1
            row += 1

        # Input fields (for form screens)
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

        # Buttons
        if screen.buttons:
            _write_section(ws, row, "操作ボタン一覧", styles)
            row += 1
            write_header_row(ws, row, ["No", "ボタン名", "アクション", "備考", "", ""], styles)
            row += 1
            for i, b in enumerate(screen.buttons):
                write_data_row(ws, row, [i+1, b.name, b.action, b.description or "", "", ""], styles)
                row += 1
            row += 1

        # Validations
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
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_screen_spec.py -v
```

Expected: 1 test PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/generators/screen_spec.py tests/test_screen_spec.py && git commit -m "feat: add screen spec sheet generator"
```

---

### Task 15: API仕様書シート生成

**Files:**
- Create: `src/generators/api_spec.py`
- Create: `tests/test_api_spec.py`

- [ ] **Step 1: テストを書く**

```python
# tests/test_api_spec.py
from openpyxl import Workbook
from src.models import ApiEndpoint, ApiParam
from src.generators.api_spec import generate_api_spec
from src.generators.styles import create_styles


def test_generate_api_spec():
    wb = Workbook()
    ws = wb.active

    endpoints = [
        ApiEndpoint(
            method="GET", url="/employee/api/search",
            controller="EmployeeController",
            request_params=[
                ApiParam(name="name", location="query", required=False, param_type="String"),
                ApiParam(name="page", location="query", required=False, param_type="int"),
            ],
            response_type="Map<String, Object>",
            description="社員検索"
        ),
        ApiEndpoint(
            method="POST", url="/employee/api/save",
            controller="EmployeeController",
            request_params=[
                ApiParam(name="employee", location="body", required=True, param_type="Employee"),
            ],
            response_type="ResponseEntity<Employee>",
            description="社員情報保存"
        ),
    ]
    styles = create_styles()
    generate_api_spec(ws, endpoints, styles)

    assert ws['A1'].value == "API仕様書"
    headers = [ws.cell(row=2, column=c).value for c in range(1, 8)]
    assert "No" in headers
    assert "メソッド" in headers
    assert "URL" in headers

    assert ws.cell(row=3, column=1).value == 1
    assert ws.cell(row=3, column=2).value == "GET"
    assert ws.cell(row=3, column=7).value == "EmployeeController"
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_api_spec.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/generators/api_spec.py
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
        data = [
            i + 1,
            desc,
            ep.method,
            ep.url,
            params_str,
            ep.response_type,
            ep.controller,
        ]
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


def _infer_description(ep: ApiEndpoint) -> str:
    """Infer a meaningful description from URL and method."""
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
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_api_spec.py -v
```

Expected: 1 test PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/generators/api_spec.py tests/test_api_spec.py && git commit -m "feat: add api spec sheet generator"
```

---

### Task 16: 業務フローシート生成

**Files:**
- Create: `src/generators/business_flow.py`
- Create: `tests/test_business_flow.py`

- [ ] **Step 1: テストを書く**

```python
# tests/test_business_flow.py
from openpyxl import Workbook
from src.generators.business_flow import generate_business_flow
from src.generators.styles import create_styles


def test_generate_business_flow():
    wb = Workbook()
    ws = wb.active

    styles = create_styles()
    generate_business_flow(ws, styles)

    assert ws['A1'].value == "業務フロー"
    # Check expense flow section
    assert "経費精算フロー" in str(ws['A3'].value)
    # Check order flow section exists
    found_order = False
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and "発注フロー" in str(cell.value):
                found_order = True
    assert found_order
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_business_flow.py -v
```

Expected: FAIL

- [ ] **Step 3: 実装**

```python
# src/generators/business_flow.py
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

    # === Expense Flow ===
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

    # Transition arrows description
    cell = ws.cell(row=row, column=1, value="状態遷移:")
    apply_header_style(cell, styles)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    row += 1

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

    # === Order Flow ===
    _write_flow_section(ws, row, "発注フロー", styles)
    row += 1

    order_states = [
        ("1", "新規", "発注の初期状態。", "発注作成時", "OrderController.save()"),
        ("2", "発注済", "仕入先に発注済み。キャンセル可能。", "発注処理実行", "（本実装では画面未実装）"),
        ("3", "納品済", "商品が納品された状態。処理完了。", "納品確認", "（本実装では画面未実装）"),
        ("4", "キャンセル", "発注がキャンセルされた状態。", "キャンセル操作", "OrderController.cancel()"),
    ]
    _write_state_table(ws, row, order_states, styles)
    row += len(order_states) + 2

    cell = ws.cell(row=row, column=1, value="状態遷移:")
    apply_header_style(cell, styles)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    row += 1

    order_transitions = [
        "新規 → 発注済 : 発注処理（未実装）",
        "発注済 → 納品済 : 納品確認（未実装）",
        "新規 → キャンセル : キャンセル操作",
        "発注済 → キャンセル : キャンセル操作（OrderController.cancel）",
        "キャンセル → 新規 : 複製操作（OrderController.copy）",
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
```

- [ ] **Step 4: テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_business_flow.py -v
```

Expected: 1 test PASS

- [ ] **Step 5: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add src/generators/business_flow.py tests/test_business_flow.py && git commit -m "feat: add business flow sheet generator"
```

---

### Task 17: メインエントリポイント + 統合テスト

**Files:**
- Create: `generate_shiyosho.py`
- Create: `tests/test_integration.py`

- [ ] **Step 1: 失敗する統合テストを書く**

```python
# tests/test_integration.py
import os
import tempfile
from generate_shiyosho import generate


def test_generate_creates_xlsx(visiondemo_path):
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = os.path.join(tmpdir, "test_output.xlsx")
        generate(visiondemo_path, output_path)

        assert os.path.exists(output_path)

        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        sheet_names = wb.sheetnames

        assert "表紙" in sheet_names
        assert "画面一覧" in sheet_names
        assert "テーブル定義" in sheet_names
        assert "ER図" in sheet_names
        assert "API仕様書" in sheet_names
        assert "業務フロー" in sheet_names

        # Check screen spec sheets exist
        screen_sheets = [s for s in sheet_names if s.startswith("SCR-")]
        assert len(screen_sheets) >= 8

        wb.close()
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_integration.py -v
```

Expected: FAIL

- [ ] **Step 3: エントリポイント実装**

```python
# generate_shiyosho.py
import os
import sys
from datetime import date
from openpyxl import Workbook
from src.project_reader import ProjectReader
from src.generators.styles import create_styles
from src.generators.cover import generate_cover
from src.generators.screen_list import generate_screen_list
from src.generators.table_def import generate_table_def
from src.generators.er_diagram import generate_er_diagram, build_relationships
from src.generators.screen_spec import generate_screen_specs
from src.generators.api_spec import generate_api_spec
from src.generators.business_flow import generate_business_flow


def generate(project_path: str, output_path: str):
    reader = ProjectReader(project_path)
    styles = create_styles()

    wb = Workbook()

    # Sheet 1: 表紙
    ws1 = wb.active
    ws1.title = "表紙"
    sys_info = reader.read_system_info()
    generate_cover(ws1, sys_info, styles)

    # Collect data
    tables = reader.read_tables()
    screens = reader.read_screens()
    endpoints, _ = reader.read_api_endpoints()
    relationships = build_relationships(tables)

    # Sheet 2: 画面一覧
    ws2 = wb.create_sheet("画面一覧")
    generate_screen_list(ws2, screens, styles)

    # Sheet 3: テーブル定義
    ws3 = wb.create_sheet("テーブル定義")
    generate_table_def(ws3, tables, styles)

    # Sheet 4: ER図
    ws4 = wb.create_sheet("ER図")
    generate_er_diagram(ws4, tables, relationships, styles)

    # Sheet 5-13: 画面仕様 (one per screen)
    generate_screen_specs(wb, screens, styles)

    # Sheet 14: API仕様書
    ws_api = wb.create_sheet("API仕様書")
    generate_api_spec(ws_api, endpoints, styles)

    # Sheet 15: 業務フロー
    ws_flow = wb.create_sheet("業務フロー")
    generate_business_flow(ws_flow, styles)

    # Save
    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    wb.save(output_path)
    print(f"式様書を生成しました: {output_path}")


def main():
    # Default paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_project = os.path.join(script_dir, '..', 'VisionDemo')
    default_output = os.path.join(default_project, 'docs', 'shiyosho',
                                  f'VisionDemo_式様書_v1.0.xlsx')

    project_path = sys.argv[1] if len(sys.argv) > 1 else default_project
    output_path = sys.argv[2] if len(sys.argv) > 2 else default_output

    project_path = os.path.abspath(project_path)
    output_path = os.path.abspath(output_path)

    if not os.path.isdir(project_path):
        print(f"エラー: プロジェクトパスが見つかりません: {project_path}")
        sys.exit(1)

    generate(project_path, output_path)


if __name__ == '__main__':
    main()
```

- [ ] **Step 4: 統合テスト実行（成功確認）**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/test_integration.py -v
```

Expected: 1 test PASS

- [ ] **Step 5: 全テスト実行**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python -m pytest tests/ -v
```

Expected: all tests PASS

- [ ] **Step 6: 本番実行**

```bash
cd /Users/mz/Documents/01_work/EasySpec && python generate_shiyosho.py
```

Expected: `VisionDemo_式様書_v1.0.xlsx` が `../VisionDemo/docs/shiyosho/` に生成される

- [ ] **Step 7: Commit**

```bash
cd /Users/mz/Documents/01_work/EasySpec && git add generate_shiyosho.py tests/test_integration.py && git commit -m "feat: add main entry point and integration test"
```

---

## まとめ

| Task | 内容 | 成果物 |
|------|------|--------|
| 1 | プロジェクト初期化 | ディレクトリ構造 + requirements.txt + conftest.py |
| 2 | データモデル定義 | `src/models.py` (9 dataclasses) |
| 3 | pom.xml パーサー | `src/parsers/pom_parser.py` |
| 4 | schema.sql パーサー | `src/parsers/schema_parser.py` |
| 5 | Entity.java パーサー | `src/parsers/entity_parser.py` |
| 6 | Controller.java パーサー | `src/parsers/controller_parser.py` |
| 7 | JSP パーサー | `src/parsers/jsp_parser.py` |
| 8 | オーケストレータ | `src/project_reader.py` |
| 9 | Excel スタイル | `src/generators/styles.py` |
| 10 | 表紙シート | `src/generators/cover.py` |
| 11 | 画面一覧シート | `src/generators/screen_list.py` |
| 12 | テーブル定義シート | `src/generators/table_def.py` |
| 13 | ER図シート | `src/generators/er_diagram.py` |
| 14 | 画面仕様シート | `src/generators/screen_spec.py` |
| 15 | API仕様書シート | `src/generators/api_spec.py` |
| 16 | 業務フローシート | `src/generators/business_flow.py` |
| 17 | エントリポイント + 統合テスト | `generate_shiyosho.py` |
