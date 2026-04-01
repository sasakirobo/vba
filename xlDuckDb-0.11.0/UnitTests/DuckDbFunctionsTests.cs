#pragma warning disable DuckDBNET001
using Xunit;
using xlDuckDb;

public class DuckDbFunctionsTests
{
    [Fact]
    public void DuckDbQuery_ReturnsExpectedRows()
    {
        var sql = "SELECT 1 AS A UNION ALL SELECT 2";
        var result = DuckDbHelper.ExecuteQuery(sql, "");
        Xunit.Assert.Equal(3, result.GetLength(0)); // header + 2 rows
        Xunit.Assert.Equal("A", result[0,0]);
        Xunit.Assert.Equal(1.0, result[1,0]);
        Xunit.Assert.Equal(2.0, result[2,0]);
    }

    [Fact]
    public void DuckDbQuery_RespectsMaxOutputRows()
    {
        xlDuckDbConfig.MaxOutputRows = 1;
        var sql = "SELECT 1 UNION ALL SELECT 2 UNION ALL SELECT 3";
        var result = DuckDbHelper.ExecuteQuery(sql, "");
        Xunit.Assert.Equal(3, result.GetLength(0)); // header + 1 row + error row
        Xunit.Assert.Contains("#ERR", result[2,0].ToString());
        xlDuckDbConfig.MaxOutputRows = 1048570; // reset
    }

    [Fact]
    public void DuckDbQueryMulti_UnionAllAndLimit()
    {
        xlDuckDbConfig.MaxOutputRows = 2;
        var sqls = new object[,] { { "SELECT 1 AS X" }, { "SELECT 2" } };
        var result = xlAddIn.DuckDbQueryMulti(sqls, "", null) as object[,];
        Xunit.Assert.NotNull(result);
        Xunit.Assert.Equal(3, result.GetLength(0)); // header + 1 row + error row
        xlDuckDbConfig.MaxOutputRows = 1048570;
    }

    [Fact]
    public void DuckDbCascadeSearch_StopsAtMaxRows()
    {
        var sqls = new object[,] { { "SELECT 1 AS X" }, { "SELECT 2" }, { "SELECT 3" } };
        var result = xlAddIn.DuckDbCascadeSearch(sqls, "", null, 2) as object[,];
        Xunit.Assert.NotNull(result);
        Xunit.Assert.Equal(4, result.GetLength(0)); // header + 2 rows + error row
        Xunit.Assert.Contains("#ERR", result[3,0].ToString());
    }

    [Fact]
    public void DuckDbNormalize_Stepwise()
    {
        var result = xlAddIn.DuckDbNormalize("㈱Test①", 511, "①", "㈱,株式会社") as object[,];
        Xunit.Assert.NotNull(result);
        Xunit.Assert.Equal("Input", result[0,0]);
        Xunit.Assert.Equal("Final", result[result.GetLength(0)-1,0]);
    }

    [Fact]
    public void DuckDbNormalize_ReturnsMultipleRows()
    {
        // すべてのオプション（511）で "㈱Test①" を正規化
        var result = xlAddIn.DuckDbNormalize("㈱Test①", 511, "①", "㈱,株式会社") as object[,];
        Xunit.Assert.NotNull(result);
        
        var rows = result.GetLength(0);
        var cols = result.GetLength(1);
        
        // 少なくとも 3 行以上（Input + 最低限のステップ + Final）
        Xunit.Assert.True(rows >= 3, $"Expected at least 3 rows, but got {rows}");
        
        // すべての行が 2 列であることを確認
        Xunit.Assert.Equal(2, cols);
        
        // 最初の行は "Input"
        Xunit.Assert.Equal("Input", result[0, 0]);
        Xunit.Assert.Equal("㈱Test①", result[0, 1]);
        
        // 最後の行は "Final"
        Xunit.Assert.Equal("Final", result[rows - 1, 0]);
    }
}
#pragma warning restore DuckDBNET001
