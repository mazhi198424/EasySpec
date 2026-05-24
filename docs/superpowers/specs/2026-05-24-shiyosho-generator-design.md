# VisionDemo 式様書 自動生成ツール — 設計仕様書

**作成日**: 2026-05-24  
**ステータス**: 設計完了 / 実装待ち  
**対象プロジェクト**: VisionDemo (Spring Boot 3.4.5 + JSP + H2)

---

## 1. 目的

VisionDemo のソースコードを静的解析し、日本の伝統的な詳細設計書（式様書）を Excel 形式で自動生成する Python スクリプトを作成する。

## 2. 技術選定

| 項目 | 決定 | 理由 |
|------|------|------|
| 言語 | Python 3 | openpyxl の Excel 制御が POI より簡潔 |
| Excel ライブラリ | openpyxl | 結合セル、罫線、フォント、列幅の精密制御が可能 |
| 実行場所 | `EasySpec/` ディレクトリ | VisionDemo 本体と分離、スクリプト単独運用 |
| 出力先 | `VisionDemo/docs/shiyosho/` | プロジェクトドキュメントと同梱 |

## 3. ファイル構成

```
EasySpec/
  generate_shiyosho.py        ← エントリポイント
  requirements.txt             ← openpyxl==3.1.5

出力:
VisionDemo/docs/shiyosho/
  VisionDemo_式様書_v1.0.xlsx
```

## 4. Excel Sheet 構成（7シート）

### 4.1 表紙
- システム名、バージョン：`pom.xml` より抽出
- 作成日：スクリプト実行日
- ドキュメント管理表（版数 / 更新日 / 更新内容 / 作成者）

### 4.2 画面一覧
全9画面の一覧表。  
データソース：Controller マッピング + JSP ファイルパス  
| 画面ID | 画面名 | URL | JSPファイル | 機能概要 |

### 4.3 テーブル定義
6テーブルの完全定義。各テーブル：
- 基本情報（テーブル名、論理名、説明）
- カラム定義表（No / 論理名 / 物理名 / 型 / 桁数 / NULL / KEY / デフォルト / 説明）
- FK 関連一覧
データソース：`schema.sql`（構造） + Entity クラス（アノテーション補足）

### 4.4 ER図
- 6テーブル間のリレーションを表形式で図示
- 各テーブルを結合セルの矩形で表現、リレーション線を罫線文字で描画
- 1:N / N:1 の多重度を明記

### 4.5 画面仕様（9画面分）
各画面につき：
- 基本情報（画面ID / URL / Controller / JSP）
- 画面レイアウト（テキスト記述）
- 表示項目一覧（リスト画面の場合、テーブルカラム定義）
- 入力項目定義（フォーム画面の場合：項目名 / 型 / 桁数 / 必須 / 入力制御 / 備考）
- 操作ボタン一覧
- バリデーションルール

### 4.6 API仕様書
全 REST API エンドポイントの一覧表（約 20 エンドポイント）。  
データソース：Controller クラスのアノテーション（`@GetMapping`, `@PostMapping`, `@PutMapping`, `@DeleteMapping`, `@RequestMapping`, `@RequestParam`, `@PathVariable`, `@RequestBody`）

| No | 機能 | メソッド | URL | リクエスト | レスポンス | Controller |
|----|------|----------|-----|-----------|-----------|------------|

### 4.7 業務フロー
2つの主要フローを状態遷移図として表形式で記述：
- **経費精算フロー**: 下書き → 申請中 → 承認済 / 差戻し
- **発注フロー**: 新規 → 発注済 → 納品済 / キャンセル

## 5. データソース解析戦略

| 情報 | ソース | 解析方法 |
|------|--------|----------|
| システム名/バージョン | `pom.xml` | XMLパース（ElementTree） |
| DBテーブル構造 | `schema.sql` | SQL 正規表現パース（CREATE TABLE + カラム定義行 + FK定義） |
| シードデータ | `data.sql` | INSERT文から初期データ値を抽出 |
| Entity補足情報 | `src/main/java/.../entity/*.java` | ファイル読み込み + `@Column` / `@JoinColumn` アノテーション正規表現抽出 |
| APIエンドポイント | `src/main/java/.../controller/*.java` | ファイル読み込み + Spring MVC アノテーション正規表現抽出 |
| 画面定義 | `src/main/webapp/WEB-INF/jsp/**/*.jsp` | ファイル読み込み + HTMLテーブル/フォーム要素抽出 |
| URL-JSPマッピング | Controller の `@GetMapping("page")` 戻り値より推測 |

## 6. Excel スタイル仕様

- **フォント**: 游ゴシックまたは MS PGothic 相当、本文 10pt、見出し 12pt bold
- **罫線**: 外枠太線、内側細線（日本式様式書標準）
- **結合セル**: タイトル行、テーブルヘッダ区切り
- **列幅**: 内容に応じて自動調整（openpyxl の `dimensions` 設定）
- **色**: ヘッダ行のみ薄い青背景（#DCE6F1）

## 7. 境界ケース

- **未実装/モック箇所**（SystemServiceの辞書、ReportServiceのランダムデータ）：`※ サンプル実装` と注記
- **ON DELETE CASCADE**: FK 定義に明記
- **@PrePersist のみで @Column 無し**: schema.sql 定義を正とし、補足として注記
- **JSP 内の JS バリデーション**: 画面仕様のバリデーション欄に記載

## 8. 含めないもの（YAGNI）

- ソースコード AST 解析（正規表現 + テキスト解析で十分）
- 動的ER図描画エンジン（固定テーブルレイアウトで描画）
- 差分・増分更新（毎回フル再生成）
- DB スキーマのリバースエンジニアリング接続
