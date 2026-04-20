# xlDuckDb 詳細マニュアル

## 概要

xlDuckDbは、Excel上でDuckDB SQLを直接実行できるアドインです。  
複数クエリのUNION ALLやカスケードサーチ、正規化ユーティリティも提供します。

---

## ⚠️ 既知の制限事項

### DuckDB.NET v1.5.0 パラメータバインディング非対応

DuckDB.NET v1.5.0 は `$1, $2, $3...` スタイルのパラメータバインディング（Prepared Statement パラメータ）をサポートしていません。

**推奨される代替パターン:**

#### 1. **xlRange テーブル関数**（最も推奨）
セル範囲を直接クエリ内で参照します：
```sql
SELECT * FROM xlRange('$A$1:$B$10') AS t(id INT, name VARCHAR)
WHERE id > 100
```

#### 2. **SQL 内に値を直接埋め込む**
Excel 関数で SQL 文字列を動的構築します：
```
=DuckDbQuery("SELECT * FROM table WHERE id = " & TEXT(A1,"0"), "db.duckdb")
```

#### 3. **文字列正規化/置換を使用**
複数の入力をまとめて処理します：
```
=DuckDbQuery("SELECT * FROM xlRange('" & A1 & "') WHERE status = 'active'", "")
```

---

## 定数・制限

- **出力最大行数（インスタンスワイド）**  
  `xlDuckDbConfig.MaxOutputRows`  
  デフォルト: 1,048,570（Excelシートの最大行数に合わせています）  
  この上限に到達した場合、出力にエラー行（`#ERR: 出力最大行数に到達`）が追加されます。

---

## 各関数の詳細

### 1. DuckDbQuery

#### 概要
DuckDB SQLを1件実行します。

#### 書式
```
=DuckDbQuery(SQL, DatabaseFile)
```

#### パラメータ
- **SQL**  
  SQL文（文字列またはセル参照）。

- **DatabaseFile**  
  DuckDBファイルパス。空欄または省略でインメモリDB。

#### 制限
- 出力最大行数は `xlDuckDbConfig.MaxOutputRows` で制限されます。

#### チュートリアル（値埋め込み）
```
=DuckDbQuery("SELECT * FROM my_table WHERE id = " & A1, "C:\\data\\test.duckdb")
```

#### チュートリアル（xlRange 使用）
```
=DuckDbQuery("SELECT * FROM xlRange('$A$1:$C$100')", "C:\\data\\test.duckdb")
```

---

### 2. DuckDbQueryMulti

#### 概要
セル範囲で複数SQLを受け取り、順次実行してUNION ALL（VSTACK）します。

#### 書式
```
=DuckDbQueryMulti(SQLRange, DatabaseFile)
```

#### パラメータ
- **SQLRange**  
  1列または1行のSQL式範囲。空セルはスキップ。  
  行数≥列数なら縦、列数>行数なら横と自動判定。

- **DatabaseFile**  
  DuckDBファイルパス。

#### 制限
- 出力最大行数は `xlDuckDbConfig.MaxOutputRows` で制限されます。
- 列数が不一致の場合は `#VALUE!` を返します。

#### チュートリアル
```
=DuckDbQueryMulti(A1:A3, "C:\\data\\test.duckdb")
```
A1:A3に複数のSQL を指定。各SQLで同じ列構造である必要があります。

---

### 3. DuckDbCascadeSearch

#### 概要
複数SQLを順次実行し、累計出力行数が`MaxRows`に到達するまでVSTACK（縦連結）します。  
指定行数に到達した時点で以降のSQLは実行しません。

#### 書式
```
=DuckDbCascadeSearch(SQLRange, DatabaseFile, Reserved, MaxRows)
```

#### パラメータ
- **SQLRange**  
  1列または1行のSQL式範囲。

- **DatabaseFile**  
  DuckDBファイルパス。

- **Reserved**  
  将来用に予約。`null` または空を指定してください。

- **MaxRows**  
  最大出力行数。

#### 制限
- 列数が不一致の場合は `#VALUE!` を返します。

#### チュートリアル
```
=DuckDbCascadeSearch(A1:A3, "C:\\data\\test.duckdb", , 1000)
```
最初のSQL が1000行出力したら、以降のSQLは実行されません。

---

### 4. DuckDbNormalize

#### 概要
文字列の正規化を適用し、最終結果のみを返します。
単一入力の場合は正規化済み文字列を返し、セル範囲入力の場合は各セルの正規化結果を配列で返します。

#### 書式
```
=DuckDbNormalize(Input, NormalizeOptions, AdditionalRemove, AdditionalReplace)
```

#### パラメータ
- **Input**  
  入力文字列またはセル範囲。
  - 単一文字列: 正規化済み文字列を返す
  - セル範囲: 各セルを正規化し、同じサイズの配列を返す（一括処理）

- **NormalizeOptions**  
  ビットフラグ（デフォルト: 511 = 全て）。
  - 1: NFKC + CaseFold
  - 2: 長音統一
  - 4: 絵文字削除
  - 8: 全角英数 → 半角
  - 16: 半角英数 → 全角
  - 32: 半角カナ → 全角
  - 64: 全角記号 → 半角
  - 128: 空白統一
  - 256: 括弧統一
  - 511: 全て

- **AdditionalRemove**  
  削除対象の文字（例: '①②③'）。

- **AdditionalReplace**  
  置換マッピング（from,to 形式。改行で複数指定。例: '㈱,株式会社'）。

#### チュートリアル（単一入力）
```
=DuckDbNormalize("Test①②③", 511, "①②③", "")
結果: "Test"
```

#### チュートリアル（セル範囲・一括処理）
```
=DuckDbNormalize(A1:A5, 511, "①②③", "㈱,株式会社")
A1:A5 の各セルを正規化し、5 つの結果を返す
```

#### チュートリアル（カスタム正規化）
```
=DuckDbNormalize(A1, 511, "①②③", "㈱,株式会社")
