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
