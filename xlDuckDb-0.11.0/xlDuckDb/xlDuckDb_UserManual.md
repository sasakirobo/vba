# xlDuckDb 詳細マニュアル

## 概要

xlDuckDbは、Excel上でDuckDB SQLを直接実行できるアドインです。  
複数クエリのUNION ALLやカスケードサーチ、正規化ユーティリティも提供します。

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
DuckDB SQLを1件実行します（パラメータ対応）。

#### 書式
```
=DuckDbQuery(SQL, DatabaseFile, [Params/Ranges...])
```

#### パラメータ
- **SQL**  
  SQL文（文字列またはセル参照）。  
  パラメータは`$1`, `$2`, ...で指定できます。

- **DatabaseFile**  
  DuckDBファイルパス。空欄または省略でインメモリDB。

- **Params/Ranges...**  
  SQLの`$1`, `$2`, ...にバインドする値やセル範囲。  
  - 文字列リテラルは `"@NORM:<flags>[;REMOVE:<chars>][;REPLACE:<from>,<to>|...]:<value>"` 形式で正規化指定可。
  - 複数セル範囲は行優先で展開されます。

#### 制限
- 出力最大行数は `xlDuckDbConfig.MaxOutputRows` で制限されます。

#### チュートリアル
```
=DuckDbQuery("SELECT * FROM my_table WHERE id = $1", "C:\\data\\test.duckdb", 123)
```

---

### 2. DuckDbQueryMulti

#### 概要
セル範囲で複数SQLを受け取り、順次実行してUNION ALL（VSTACK）します。

#### 書式
```
=DuckDbQueryMulti(SQLRange, DatabaseFile, [ArgsRange])
```

#### パラメータ
- **SQLRange**  
  1列または1行のSQL式範囲。空セルはスキップ。  
  行数≥列数なら縦、列数>行数なら横と自動判定。

- **DatabaseFile**  
  DuckDBファイルパス。

- **ArgsRange**  
  各SQLのパラメータ範囲（省略可）。  
  行=SQL番号、列=$1,$2,...  
  末尾の空セルは自動トリム。

#### 制限
- 出力最大行数は `xlDuckDbConfig.MaxOutputRows` で制限されます。
- 列数が不一致の場合は `#VALUE!` を返します。

#### チュートリアル
```
=DuckDbQueryMulti(A1:A3, "C:\\data\\test.duckdb", B1:C3)
```
A1:A3にSQL、B1:C3に各SQLのパラメータを指定。

---

### 3. DuckDbCascadeSearch

#### 概要
複数SQLを順次実行し、累計出力行数が`MaxRows`に到達するまでVSTACK（縦連結）します。  
指定行数に到達した時点で以降のSQLは実行しません。

#### 書式
```
=DuckDbCascadeSearch(SQLRange, DatabaseFile, ArgsRange, MaxRows)
```

#### パラメータ
- **SQLRange**  
  1列または1行のSQL式範囲。

- **DatabaseFile**  
  DuckDBファイルパス。

- **ArgsRange**  
  各SQLのパラメータ範囲。

- **MaxRows**  
  累計出力行数の上限（整数）。

#### 制限
- `MaxRows`に到達した場合、エラー行（`#ERR: 累計出力最大行数に到達`）が追加されます。
- 列数が不一致の場合は `#VALUE!` を返します。

#### チュートリアル
```
=DuckDbCascadeSearch(A1:A10, "C:\\data\\test.duckdb", B1:C10, 100)
```
A1:A10にSQL、B1:C10に各SQLのパラメータ、最大100行まで出力。

---

### 4. DuckDbNormalize

#### 概要
文字列正規化の各ステップを確認できます。

#### 書式
```
=DuckDbNormalize(Input, [NormalizeOptions], [AdditionalRemove], [AdditionalReplace])
```

#### パラメータ
- **Input**  
  入力文字列またはセル範囲。

- **NormalizeOptions**  
  ビットフラグ（省略時: 511=全て適用）。  
  - 1=NFKC+CaseFold
  - 2=LongVowel
  - 4=EmojiRemove
  - 8=FullwidthAlnum
  - 16=HalfAlnum
  - 32=HalfKana
  - 64=FullSymbol
  - 128=Whitespace
  - 256=Bracket

- **AdditionalRemove**  
  除去する追加文字。

- **AdditionalReplace**  
  置換マッピング（1行1組、カンマ区切り）。

#### チュートリアル
```
=DuckDbNormalize("㈱Test①", 511, "①", "㈱,株式会社")
```
→ 正規化ステップごとの変化を2列で出力。

---

## パラメータ指定のヒント

- SQLやパラメータはセル参照でも直接入力でもOK。
- パラメータ範囲は「行=SQL番号、列=パラメータ番号」で指定。
- 文字列パラメータは `"@NORM:..."` 形式で正規化指定が可能。

---

## よくある質問

- **Q. 出力が多すぎてExcelが固まるのを防げますか？**  
  A. `xlDuckDbConfig.MaxOutputRows` で全関数の出力最大行数を制限できます。

- **Q. SQLのパラメータはどう指定する？**  
  A. `$1`, `$2`, ...でSQL内に記述し、関数の引数またはパラメータ範囲で値を渡します。

---

## 参考

- [DuckDB公式ドキュメント](https://duckdb.org/docs/index)
- [ExcelDna公式](https://excel-dna.net/)

---
