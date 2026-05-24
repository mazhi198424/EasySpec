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

        entity_dir = self._find_package_dir('entity')
        entity_infos = {}
        if entity_dir:
            for java_file in glob.glob(os.path.join(entity_dir, '*.java')):
                entity_info = parse_entity(java_file)
                entity_infos[entity_info['table_name']] = entity_info

        for table in tables:
            if table.name in entity_infos:
                ei = entity_infos[table.name]
                self._merge_entity_info(table, ei)

        self._assign_logical_names(tables)
        return tables

    def _merge_entity_info(self, table, entity_info):
        for col in table.columns:
            for ei_col in entity_info['columns']:
                if ei_col['name'] == col.name:
                    if ei_col.get('length') and not col.length:
                        col.length = ei_col['length']
                    if 'nullable' in ei_col and ei_col['nullable'] is False:
                        col.nullable = False
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
            jsp_path = os.path.join(self.webapp, jsp)
            if os.path.exists(jsp_path):
                parsed = parse_jsp(jsp_path)
                screen.fields = parsed['fields']
                screen.buttons = parsed['buttons']
                screen.table_columns = parsed['table_columns']
                if parsed['layout_description']:
                    screen.layout_description += ' | ' + parsed['layout_description']
            if ctrl == 'ReportController' or ctrl == 'SystemController':
                screen.is_mock = True
            screens.append(screen)

        return screens

    def _find_package_dir(self, subpackage: str) -> str:
        base = self.java_src
        if not os.path.isdir(base):
            return None
        for root, dirs, files in os.walk(base):
            if os.path.basename(root) == subpackage:
                return root
        return None
