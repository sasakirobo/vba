# xlDuckDb ユーザーガイド

> Excel から DuckDB を直接クエリできる Excel アドイン。  
> 高速な分析クエリ・Prepared Statement・LRU キャッシュ・並列実行に対応。

---

## 目次

1. [はじめに](#1-はじめに)
2. [関数リファレンス — DuckDbQuery](#2-関数リファレンス--duckdbquery)
3. [パラメータ指定方法](#3-パラメータ指定方法)
4. [xlRange — Excelセル範囲をテーブルとして使う](#4-xlrange--excelセル範囲をテーブルとして使う)
5. [キャッシュ機構の解説](#5-キャッシュ機構の解説)
6. [接続プールの解説](#6-接続プールの解説)
7. [エラー値リファレンス](#7-エラー値リファレンス)
8. [チュートリアル](#8-チュートリアル)
9. [制限事項と注意点](#9-制限事項と注意点)

---

## 1. はじめに

xlDuckDb は [DuckDB](https://duckdb.org/) を Excel から直接利用するためのアドインです。  
`DuckDbQuery` 関数 1 つで SQL クエリを実行し、結果をセル範囲にスピルさせることができます。

### 主な特徴

- **Prepared Statement** — SQL を事前コンパイルしてバインド変数で安全・高速に実行
- **LRU キャッシュ** — Prepared Statement と実行結果を自動キャッシュ
- **接続プール** — 接続を使い回してオープン/クローズのコストをゼロに
- **並列実行対応** — Excel のマルチスレッド再計算（`IsThreadSafe`）に対応
- **xlRange** — Excel のセル範囲を DuckDB のテーブルとして直接クエリ
- **インメモリ DB** — ファイルなしでその場限りの DuckDB データベースを利用可能

---

## 2. 関数リファレンス — DuckDbQuery

```
=DuckDbQuery(SQL, [DatabaseFile], [Param1/Range1], [Param2/Range2], ...)
```

### 引数

| 引数 | 型 | 必須 | 説明 |
|---|---|---|---|
| `SQL` | 文字列またはセル参照 | ✅ | 実行する SQL 文。`$1`, `$2`, ... でパラメータを埋め込める。数式が入ったセルを参照すると、その計算結果の文字列が SQL として使われる。 |
| `DatabaseFile` | 文字列 | — | DuckDB ファイルのパス（例: `C:\data\mydb.duckdb`）。省略または空文字でインメモリ DB を使用。 |
| `Param1`, `Param2`, ... | セル / セル範囲 / リテラル値 | — | SQL の `$1`, `$2`, ... にバインドする値。詳細は [§3](#3-パラメータ指定方法) を参照。 |

### 戻り値

クエリ結果を 2D 配列（スピル範囲）として返します。  
1 行目は**列名（ヘッダー）**、2 行目以降がデータ行です。

| 状況 | 戻り値 |
|---|---|
| 結果あり | ヘッダー行 + データ行の 2D 配列 |
| 結果なし（0件） | `#N/A` |
| SQL が空 / 参照先セルが空 | `#VALUE!` |
| 実行エラー | `#ERR - <エラーメッセージ>` |

---

## 3. パラメータ指定方法

SQL 内に `$1`, `$2`, `$3`, ... と書いた箇所に、第3引数以降の値が**順番に**バインドされます。

### 3-1. リテラル値を渡す

```excel
=DuckDbQuery("SELECT * FROM sales WHERE year=$1 AND region=$2", "", 2024, "East")
```

`$1` ← `2024`、`$2` ← `"East"` がバインドされます。

### 3-2. 単一セルを渡す

```excel
=DuckDbQuery("SELECT * FROM sales WHERE year=$1", "", A1)
```

A1 セルの**計算済みの値**が `$1` にバインドされます。A1 に数式が入っていても、その結果が使われます。

### 3-3. セル範囲を渡す（行優先展開）

複数セルの範囲を渡すと、**行優先**で 1 セルずつ展開されて `$1`, `$2`, ... にバインドされます。

```excel
=DuckDbQuery("INSERT INTO t VALUES ($1, $2, $3)", "", A1:C1)
```

| セル | バインド先 |
|---|---|
| A1 | `$1` |
| B1 | `$2` |
| C1 | `$3` |

複数行の範囲も同様です：

```excel
=DuckDbQuery("SELECT * FROM t WHERE (a,b) IN (($1,$2), ($3,$4))", "", A1:B2)
```

| セル | バインド先 |
|---|---|
| A1 | `$1` |
| B1 | `$2` |
| A2 | `$3` |
| B2 | `$4` |

### 3-4. 複数の引数を組み合わせる

```excel
=DuckDbQuery("SELECT * FROM t WHERE year=$1 AND name=$2 AND score>=$3", "", A1, B1, C1)
```

### 3-5. SQL をセルに書いて参照する

長い SQL はセル（例: `A1`）に書いておき、参照で渡せます。

```excel
' A1 の内容: SELECT name, SUM(amount) FROM sales WHERE year=$1 GROUP BY name
=DuckDbQuery(A1, "", B1)
```

### 3-6. 空セル・エラーセルのバインド

| セルの状態 | バインドされる値 |
|---|---|
| 空セル | `NULL` |
| エラー値（`#N/A` など） | `NULL` |
| 数値 | 数値（`double`） |
| 文字列 | 文字列 |
| 論理値 | `true` / `false` |

---

## 4. xlRange — Excelセル範囲をテーブルとして使う

SQL 内に `xlRange` キーワードを含めると、Excel のセル範囲を **DuckDB のテーブル**として `SELECT` できます。

### 制約

- **1行目**: 列名（ヘッダー）
- **2行目**: データ型の推論行（数値セルは `DOUBLE`、文字列は `VARCHAR`、論理値は `BOOLEAN`）
- **3行目以降**: データ

### 単一範囲

```excel
=DuckDbQuery("SELECT name, SUM(amount) FROM xlRange GROUP BY name", "", A1:B10)
```

A1:B10 が `xlRange` テーブルとして DuckDB に渡されます。

### 複数範囲

複数範囲を使う場合は `xlRange[1]`, `xlRange[2]`, ... と書きます。

```excel
=DuckDbQuery(
  "SELECT a.name, b.score FROM xlRange[1] a JOIN xlRange[2] b ON a.id = b.id",
  "",
  A1:B20,
  D1:E20
)
```

### xlRange と $N パラメータの併用

xlRange テーブルと `$N` スカラーパラメータは同時に使えます。  
`xlRange` を含む複数セル参照は xlRange テーブルとして扱われ、それ以外の引数はスカラーパラメータとして `$1`, `$2`, ... に展開されます。

```excel
=DuckDbQuery(
  "SELECT * FROM xlRange WHERE amount > $1",
  "",
  A1:C20,
  D1
)
```

- `A1:C20` → `xlRange` テーブル
- `D1`（単一セル）→ `$1`

> **注意**: `xlRange` を使うクエリは結果キャッシュの対象外です（セルの値が変化するたびに再実行されます）。

---

## 5. キャッシュ機構の解説

xlDuckDb は 2 層のキャッシュを持ちます。いずれも **LRU（Least Recently Used）** アルゴリズムで管理され、上限を超えると最も古いエントリが自動的に削除されます。

### 5-1. Prepared Statement キャッシュ

SQL を DuckDB に送信する前に**事前コンパイル（Prepare）**し、その結果をキャッシュします。

```
初回実行:  SQL テキスト → DuckDB がパース・最適化・コンパイル → キャッシュに保存
2回目以降: キャッシュから取得 → パラメータをバインドして即実行（パース不要）
```

| 項目 | 値 |
|---|---|
| キャッシュキー | `(接続インスタンス, SQL 文字列)` |
| 上限件数 | **256 件** |
| 追い出し戦略 | LRU（最も長く使われていない Statement を削除） |
| 接続再作成時 | キャッシュ全クリア（古い Statement の混在を防ぐ） |

**効果**: 同じ SQL を繰り返し実行するケース（パラメータだけ変えて実行するバッチ処理など）で特に有効です。

### 5-2. 実行結果キャッシュ

クエリの**実行結果**（2D 配列）をそのままキャッシュします。

```
初回実行:  SQL + パラメータ値 → DuckDB 実行 → 結果をキャッシュに保存
2回目以降: キャッシュキーが一致 → DuckDB を呼ばずに即返却
```

| 項目 | 値 |
|---|---|
| キャッシュキー | `(データソース, SQL 文字列, パラメータ値の配列)` |
| 上限件数 | **128 件** |
| 追い出し戦略 | LRU |
| xlRange クエリ | **対象外**（セルの値が変化するたびに結果が変わるため） |

**効果**: 全く同じ SQL とパラメータで再計算が走った場合（Excel の再描画や他セルの更新など）に、DB へのアクセスをスキップして即座に結果を返します。

### 5-3. LRU アルゴリズムの動作イメージ

```
キャパシティ: 3 件の例

追加 A → [A]
追加 B → [B, A]
追加 C → [C, B, A]
取得 A → [A, C, B]   ← A が先頭に移動
追加 D → [D, A, C]   ← 最古の B が追い出される
```

上限に達しても**よく使うエントリは保持**され、滅多に使わないエントリだけが削除されます。

### 5-4. キャッシュが有効でないケース

以下の場合、結果キャッシュはヒットせず毎回 DuckDB が実行されます：

- `xlRange` を使っているクエリ
- パラメータ値が異なる（セルの値が変わった）
- SQL 文字列が異なる（大文字小文字も含め完全一致が必要）
- データソース（DB ファイルパス）が異なる

---

## 6. 接続プールの解説

DuckDB の接続（`DuckDBConnection`）はオープン/クローズのコストが高いため、xlDuckDb は接続を**プールして使い回し**ます。

### プールの動作

```
1. DuckDbQuery が呼ばれる
2. プールからアイドル状態の接続を取得（なければ新規作成）
3. クエリを実行
4. 実行完了後、接続をプールに返却（クローズしない）
5. 次の DuckDbQuery 呼び出しで同じ接続を再利用
```

### プールサイズ

Excel の並列再計算スレッド数に合わせて自動決定されます。

```
プールサイズ = 論理コア数 × 2  （最小: 4、最大: 64）
```

例: 8コア PC → プールサイズ = 16

環境変数 `XLDUCKDB_POOL_SIZE` で上書きできます：

```
setx XLDUCKDB_POOL_SIZE 32
```

### データソースごとの独立

DB ファイルパスが異なる場合、それぞれ独立したプールが作られます。

```
:memory:              → プールA（最大 N 接続）
C:\data\sales.duckdb  → プールB（最大 N 接続）
C:\data\hr.duckdb     → プールC（最大 N 接続）
```

### タイムアウト

全接続が使用中で 30 秒以内に空きが出ない場合、`#ERR` を返します。  
頻発する場合は `XLDUCKDB_POOL_SIZE` を増やしてください。

---

## 7. エラー値リファレンス

| 表示 | 意味 | 対処法 |
|---|---|---|
| `#N/A` | クエリ結果が 0 件 | SQL や条件を確認する |
| `#VALUE!` | SQL が空、またはセル参照先が空/エラー | SQL 引数を確認する |
| `#NULL!` | 関数ウィザード表示中（正常） | 数式確定後に消える |
| `#ERR - <メッセージ>` | DuckDB 実行エラーまたは接続エラー | メッセージを確認。SQL 構文やファイルパスをチェック |

---

## 8. チュートリアル

### チュートリアル 1: インメモリ DB で集計する

最もシンプルな使い方です。DuckDB のインメモリ DB に対してクエリを実行します。

**セルの準備**

| セル | 内容 |
|---|---|
| A1 | `=DuckDbQuery("SELECT 42 AS answer, 'hello' AS greeting", "")` |

**結果**（A1 からスピル）

| answer | greeting |
|---|---|
| 42 | hello |

---

### チュートリアル 2: CSV / Parquet ファイルをクエリする

DuckDB はファイルを直接クエリできます。DB ファイルは不要です。

```excel
=DuckDbQuery("SELECT region, SUM(sales) AS total FROM 'C:\data\sales.csv' GROUP BY region ORDER BY total DESC", "")
```

---

### チュートリアル 3: パラメータでフィルタリングする

**シート構成**

| セル | 内容 | 説明 |
|---|---|---|
| A1 | `2024` | 年（パラメータ） |
| B1 | `East` | 地域（パラメータ） |
| A3 | ↓ 数式 | 結果出力先 |

**数式（A3 に入力）**

```excel
=DuckDbQuery(
  "SELECT name, amount FROM 'C:\data\sales.parquet' WHERE year=$1 AND region=$2 ORDER BY amount DESC",
  "",
  A1,
  B1
)
```

A1 や B1 の値を変えると自動再計算されます。

---

### チュートリアル 4: SQL をセルに書いて管理する

長い SQL はセルに書いておくと管理しやすくなります。

**シート構成**

| セル | 内容 |
|---|---|
| B1 | `SELECT name, SUM(amount) AS total FROM 'C:\data\sales.parquet' WHERE year=$1 GROUP BY name HAVING total > $2 ORDER BY total DESC` |
| B2 | `2024`（年） |
| B3 | `100000`（最小金額） |

**数式（A5 に入力）**

```excel
=DuckDbQuery(B1, "", B2, B3)
```

B1 の SQL を直接編集するだけでクエリを変更できます。

---

### チュートリアル 5: Excel セル範囲を DuckDB でクエリする（xlRange）

Excel 上のデータをそのまま DuckDB で集計します。

**シート上のデータ（A1:C6）**

| name | amount | region |
|---|---|---|
| string | number | string |
| Alice | 120 | East |
| Bob | 85 | West |
| Carol | 200 | East |
| Dave | 60 | West |

> 1 行目: ヘッダー（列名）  
> 2 行目: 型推論行（`string` → VARCHAR, `number` → DOUBLE）  
> 3 行目以降: データ

**数式（E1 に入力）**

```excel
=DuckDbQuery(
  "SELECT region, SUM(amount) AS total FROM xlRange GROUP BY region ORDER BY total DESC",
  "",
  A1:C6
)
```

**結果**

| region | total |
|---|---|
| East | 320 |
| West | 145 |

---

### チュートリアル 6: 複数テーブルを JOIN する（xlRange 複数）

**シート構成**

| 範囲 | 内容 |
|---|---|
| A1:B5 | 社員マスタ（id, name） |
| D1:E5 | スコアデータ（id, score） |

**数式**

```excel
=DuckDbQuery(
  "SELECT a.name, b.score FROM xlRange[1] a JOIN xlRange[2] b ON a.id = b.id ORDER BY b.score DESC",
  "",
  A1:B5,
  D1:E5
)
```

---

### チュートリアル 7: xlRange とパラメータを組み合わせる

xlRange テーブルを使いながら、条件をセル参照で渡せます。

**シート構成**

| セル | 内容 |
|---|---|
| A1:C20 | 売上データ（xlRange として渡す） |
| E1 | `100`（閾値） |

**数式**

```excel
=DuckDbQuery(
  "SELECT name, amount FROM xlRange WHERE amount > $1 ORDER BY amount DESC",
  "",
  A1:C20,
  E1
)
```

- `A1:C20` → `xlRange` テーブル（SQL に `xlRange` が含まれるため）
- `E1`（単一セル）→ `$1` にバインド

---

### チュートリアル 8: ファイル DB を使って結果を永続化する

インメモリ DB と違い、DuckDB ファイルに書き込んで永続化できます。

```excel
' テーブル作成（初回のみ）
=DuckDbQuery("CREATE TABLE IF NOT EXISTS kpi (month TEXT, value DOUBLE)", "C:\data\report.duckdb")

' データ挿入
=DuckDbQuery("INSERT INTO kpi VALUES ($1, $2)", "C:\data\report.duckdb", A1, B1)

' 集計クエリ
=DuckDbQuery("SELECT month, value FROM kpi ORDER BY month", "C:\data\report.duckdb")
```

---

## 9. 制限事項と注意点

### xlRange の制限

- xlRange を使うクエリは**結果キャッシュの対象外**です。セルが変化するたびに DuckDB が再実行されます。
- xlRange の **1 行目は列名**、**2 行目は型推論行**（データとしてカウントされません）。データは 3 行目以降から始まります。
- xlRange に空セルや非対応の型がある場合、`NULL` または `VARCHAR` として扱われます。

### 並列実行

- `IsThreadSafe = true` が設定されており、Excel のマルチスレッド再計算に対応しています。
- xlRange を使う場合、セルデータの読み取りは Excel メインスレッド上で先行して完了させた後、DuckDB ワーカースレッドに渡されます。

### DB ファイルの排他制御

- DuckDB のファイル DB は**書き込み時に排他ロック**がかかります。別のプロセスが同じファイルを開いている場合、接続エラーになります。
- 読み取り専用クエリ（`SELECT`）のみであれば複数接続から同時にアクセスできます。

### キャッシュの注意

- 結果キャッシュは SQL 文字列の**完全一致**（大文字小文字を含む）でキーを判定します。`SELECT` と `select` は別エントリになります。
- Excel アドインを無効化（`AutoClose`）すると全キャッシュと全接続が解放されます。
