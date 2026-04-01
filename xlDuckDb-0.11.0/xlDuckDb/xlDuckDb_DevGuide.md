# xlDuckDb 開発者向け仕様・UnitTest解説

## 1. 概要

xlDuckDbはExcel-DNAを利用したDuckDBアドインです。C#/.NET 8で実装されています。

- SQL実行関数（単一/複数/カスケード）
- 文字列正規化関数
- インスタンスワイドな出力最大行数制限

## 2. 依存ライブラリ

- DuckDB.NET.Data
- ExcelDna.Integration
- .NET 8 標準ライブラリ

### 注意点
- Excel-DNAのバージョンとDuckDB.NETのバージョン互換性に注意。
- Excelのセル範囲・型変換はExcelDna.IntegrationのAPIを利用。
- 出力最大行数はExcelの仕様（1048576行）を超えないようにする。

## 3. 機能拡張時の注意

- 新関数追加時は、出力最大行数制限（xlDuckDbConfig.MaxOutputRows）を必ず考慮。
- パラメータ展開・型変換は既存ヘルパー（NormParam, ExcelHelper等）を流用。
- エラー時はExcelErrorまたはエラー行で返す。

## 4. UnitTest仕様

- MSTest/xUnit/NUnitいずれも可（例はxUnit）。
- DuckDBのインメモリDBを使い、ファイルI/O不要。
- 各関数ごとに正常系・異常系・最大行数制限系をテスト。
- テストデータはテーブル作成・データ挿入SQLでセットアップ。

### サンプルUnitTest（xUnit例）

```csharp
using Xunit;
using xlDuckDb;

public class DuckDbFunctionsTests
{
    [Fact]
    public void DuckDbQuery_ReturnsExpectedRows()
    {
        var sql = "SELECT 1 AS A UNION ALL SELECT 2";
        var result = DuckDbHelper.ExecuteQuery(sql, "");
        Assert.Equal(3, result.GetLength(0)); // header + 2 rows
        Assert.Equal("A", result[0,0]);
        Assert.Equal(1.0, result[1,0]);
        Assert.Equal(2.0, result[2,0]);
    }

    [Fact]
    public void DuckDbQuery_RespectsMaxOutputRows()
    {
        xlDuckDbConfig.MaxOutputRows = 1;
        var sql = "SELECT 1 UNION ALL SELECT 2 UNION ALL SELECT 3";
        var result = DuckDbHelper.ExecuteQuery(sql, "");
        Assert.Equal(3, result.GetLength(0)); // header + 1 row + error row
        Assert.Contains("#ERR", result[2,0].ToString());
        xlDuckDbConfig.MaxOutputRows = 1048570; // reset
    }

    [Fact]
    public void DuckDbQueryMulti_UnionAllAndLimit()
    {
        xlDuckDbConfig.MaxOutputRows = 2;
        var sqls = new object[,] { { "SELECT 1 AS X" }, { "SELECT 2" } };
        var result = xlAddIn.DuckDbQueryMulti(sqls, "", null) as object[,];
        Assert.NotNull(result);
        Assert.Equal(3, result.GetLength(0)); // header + 1 row + error row
        xlDuckDbConfig.MaxOutputRows = 1048570;
    }
}
```

## 5. テスト拡張例
- パラメータ展開、正規化、カスケードサーチの停止条件、エラー系も網羅すること。
- テストごとにMaxOutputRowsをリセットすること。

---
