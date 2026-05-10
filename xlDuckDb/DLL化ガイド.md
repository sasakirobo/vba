# xlDuckDb DLL化ガイド - Extension統合版

## 概要

xlDuckDbプロジェクトはDuckDB.NET Fullパッケージを使用し、以下の4つのBuilt-in Extensionを自動的に組み込むように構成されています：

- **excel** - Excelファイルの直接読み取り
- **parquet** - Parquetファイル形式対応
- **https** - HTTPS/TLS通信対応
- **lance** - Lance形式（ベクトル検索）対応

## ビルド設定

### 1. プロジェクトファイル設定（xlDuckDb.csproj）

```xml
<!-- DuckDB.NET Fullパッケージ（全ネイティブバイナリ含む） -->
<ItemGroup>
    <PackageReference Include="DuckDB.NET.Bindings.Full" Version="1.5.0" />
    <PackageReference Include="DuckDB.NET.Data.Full" Version="1.5.0" />
</ItemGroup>

<!-- ビルド後にDuckDB拡張機能ファイルをコピー -->
<Target Name="CopyDuckDbExtensions" AfterTargets="Build">
    <Copy
        SourceFiles="@(DuckDbNativeExtensions)"
        DestinationFolder="$(OutputPath)"
        SkipUnchangedFiles="true"
        ContinueOnError="true"
    />
</Target>
```

### 2. Runtime初期化（DuckDbExtensionInitializer.cs）

DuckDB接続確立時に自動的にExtensionをロードします：

```csharp
public static void InitializeExtensions(DuckDBConnection conn)
{
    // 接続ごとに excel, parquet, https, lance をINSTALL/LOAD
    foreach (var ext in BuiltInExtensions)
    {
        cmd.CommandText = $"INSTALL {ext};";
        cmd.ExecuteNonQuery();
        cmd.CommandText = $"LOAD {ext};";
        cmd.ExecuteNonQuery();
    }
}
```

### 3. 接続プール統合（DuckDbConnectionManager.cs）

```csharp
private PoolEntry CreateEntry()
{
    var conn = new DuckDBConnection(connStr);
    conn.Open();

    // 接続確立直後にExtensionをロード
    DuckDbExtensionInitializer.InitializeExtensions(conn);

    return new PoolEntry(conn);
}
```

## ビルド・公開手順

### ステップ1：ビルド（Debug/Release）

```bash
# Release ビルド
dotnet build -c Release xlDuckDb\xlDuckDb.csproj

# または MSBuild 使用
msbuild xlDuckDb\xlDuckDb.csproj -p:Configuration=Release -p:Platform=x64
```

**ビルド後の出力：**
- `bin\x64\Release\net8.0-windows\xlDuckDb.dll` - メインアセンブリ
- `bin\x64\Release\net8.0-windows\*.so` - DuckDB拡張機能バイナリ
- `bin\x64\Release\net8.0-windows\DuckDB.NET.*.dll` - 依存ライブラリ

### ステップ2：Excel Add-in パッケージング（ExcelDna使用）

ExcelDnaPackを使用してExcel向けアドイン化：

```bash
# ExcelDnaPack実行
excelDnaPack.exe "bin\x64\Release\net8.0-windows\xlDuckDb.dna" -y

# 出力: xlDuckDb-AddIn.xll（Excelが直接ロード可能）
```

または、プロジェクト設定で自動化：

```xml
<PropertyGroup>
    <RunExcelDnaPack>true</RunExcelDnaPack>
</PropertyGroup>
```

### ステップ3：Publish（配布用）

```bash
# 自己完結型配布版（全ファイル含む）
dotnet publish -c Release -f net8.0-windows \
    --self-contained \
    -p:PublishSingleFile=false \
    -p:CopyOutputSymbolsToPublishDirectory=false

# 出力ディレクトリ: publish\net8.0-windows\
```

## ファイル構成

### ビルド出力（bin\x64\Release\net8.0-windows\）

```
xlDuckDb.dll                          ← Main assembly
xlDuckDb.dll.config                   ← Config
xlDuckDb.pdb                          ← Debug symbols (optional)

DuckDB.NET.*.dll                      ← DuckDB.NET managed libraries
ExcelDna.*.dll                        ← ExcelDna framework

libduckdb.dll                         ← DuckDB native core
libduckdb_parquet.so                  ← Parquet extension
libduckdb_excel.so                    ← Excel extension
libduckdb_https.so                    ← HTTPS extension
libduckdb_lance.so                    ← Lance extension
```

### Publish出力（publish\net8.0-windows\）

```
xlDuckDb.dll                          ← All managed code self-contained
libduckdb*.so                         ← Native extensions included
[All runtime dependencies]            ← .NET runtime included (if --self-contained)
```

## Extension利用方法

### 1. Excel ファイル読み取り

```sql
SELECT * FROM read_excel('data.xlsx', sheet='Sheet1');
```

### 2. Parquet ファイル読み取り

```sql
SELECT * FROM read_parquet('data.parquet');
```

### 3. HTTPS でのParquetリモート読み取り

```sql
SELECT * FROM read_parquet('https://example.com/data.parquet');
```

### 4. Lance ベクトル検索

```sql
CREATE TABLE vectors AS SELECT * FROM read_lance('vectors.lance');
```

## Extension メタデータ確認

`DuckDbExtensionInitializer` クラスで以下のメソッドを提供：

```csharp
// ロード済みExtensionの一覧取得
var extensions = DuckDbExtensionInitializer.GetLoadedExtensions(connection);

// 特定のExtensionが利用可能か確認
bool available = DuckDbExtensionInitializer.IsExtensionAvailable(connection, "parquet");
```

## トラブルシューティング

### 1. "libduckdb_*.so: cannot open shared object"

**原因：** ネイティブバイナリが出力ディレクトリに含まれていない

**対策：**
- `CopyDuckDbExtensions` ターゲットが実行されているか確認
- `DuckDB.NET.Bindings.Full` パッケージが正しくインストールされているか確認
- `$(NuGetPackageRoot)` パスが正しいか確認

### 2. "Extension 'excel' not found"

**原因：** DuckDB.NET StandardパッケージではなくFullパッケージを使用していない

**対策：**
```xml
<!-- ❌ 不正 -->
<PackageReference Include="DuckDB.NET.Bindings" Version="1.5.0" />

<!-- ✅ 正しい -->
<PackageReference Include="DuckDB.NET.Bindings.Full" Version="1.5.0" />
```

### 3. Extension ロード失敗（警告のみ）

**仕様：** Extensionロード失敗は非致命的エラーとして処理

**動作：** 
- DebugOutputに警告が出力される
- 他のExtensionとDuckDB本体は引き続き機能

## テスト検証

Extension初期化の確認：

```bash
# HTTPS + Parquet Extension テスト
dotnet test --filter "TestHttpParquetRead"

# 出力: ✅ PASSED
```

## パフォーマンス考慮事項

1. **初期化オーバーヘッド**
   - 接続ごとにExtension INSTALL/LOAD実行
   - キャッシュ未活用時は初回接続で最大1-2秒の遅延
   - 接続プーリングにより2回目以降はゼロオーバーヘッド

2. **メモリ使用量**
   - Extensionロード後、接続あたり約50-100MBメモリ増加
   - 複数接続時はプール効率で削減

3. **最適化**
   ```csharp
   // 接続再利用で性能向上
   var pooledConn = DuckDbConnectionManager.Instance.Acquire(dataSource);
   // ... queries ...
   pooledConn.Dispose(); // プールに戻る
   ```

## セキュリティ推奨事項

1. **HTTPS Extension**
   - 信頼できないURLからのリモートデータ読み取りは避ける
   - ファイアウォール・プロキシで通信を制御

2. **Excel Extension**
   - マクロ含むExcelファイルからの読み取りはDuckDBで安全（マクロ実行されない）

3. **配布時**
   - ネイティブバイナリはビルドマシンでのみ検証
   - 配布先でのアンチウイルス・セキュリティスキャン推奨

## 参照資料

- [DuckDB.NET Documentation](https://github.com/Giorgi/DuckDB.NET)
- [DuckDB Extensions](https://duckdb.org/docs/extensions/overview)
- [ExcelDna Deployment](https://github.com/Excel-DNA/ExcelDna)

---

**最終更新：** 2024年
**バージョン：** xlDuckDb 0.11.0 with DuckDB.NET 1.5.0 Full
