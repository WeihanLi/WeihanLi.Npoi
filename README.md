# WeihanLi.Npoi

[![WeihanLi.Npoi](https://img.shields.io/nuget/v/WeihanLi.Npoi)](https://www.nuget.org/packages/WeihanLi.Npoi/) [![WeihanLi.Npoi Latest](https://img.shields.io/nuget/vpre/WeihanLi.Npoi)](https://www.nuget.org/packages/WeihanLi.Npoi/absoluteLatest) [![NuGet Downloads](https://img.shields.io/nuget/dt/WeihanLi.Npoi)](https://www.nuget.org/packages/WeihanLi.Npoi/)

## Build Status

[![Azure Pipeline Build Status](https://weihanli.visualstudio.com/Pipelines/_apis/build/status/WeihanLi.WeihanLi.Npoi?branchName=dev)](https://weihanli.visualstudio.com/Pipelines/_build/latest?definitionId=13&branchName=dev) [![Github Build Status](https://github.com/WeihanLi/WeihanLi.Npoi/actions/workflows/dotnet.yml/badge.svg)](https://github.com/WeihanLi/WeihanLi.Npoi/actions/workflows/dotnet.yml)

## Introduction

[NPOI](https://github.com/tonyqus/npoi) extensions based on target framework `netstandard2.0`.

WeihanLi.Npoi provides a powerful and easy-to-use toolkit for working with Excel and CSV files in .NET applications. It offers:

- **Simple API**: Intuitive extension methods for common import/export operations
- **Flexible Configuration**: Support for both Attribute-based and FluentAPI configuration
- **High Performance**: Optimized for handling large datasets efficiently
- **Rich Features**: Advanced capabilities like template export, multi-sheet support, and shadow properties
- **CSV Support**: Full support for CSV file operations alongside Excel

## Core Features

### 📥 Data Import

- Import Excel files to `List<TEntity>` or `IEnumerable<TEntity>`
- Import Excel files to `DataTable`
- Import CSV files to entities or DataTable
- Support for custom header rows and sheet selection
- Automatic type conversion and data mapping

### 📤 Data Export

- Export `IEnumerable<TEntity>` or `DataTable` to Excel files (.xls/.xlsx)
- Export data to Excel byte arrays or streams
- Export to CSV files or byte arrays
- Template-based export with placeholders for complex layouts
- Multi-sheet export in a single workbook

### ⚙️ Configuration Options

- **Attribute Configuration**: Simple decoration with `[Column]` and `[Sheet]` attributes
- **FluentAPI Configuration**: Powerful and flexible configuration with fluent syntax (Recommended)
- Custom column mapping, formatting, and transformations
- Support for shadow properties (columns not in the model)

### 🎨 Advanced Capabilities

- **InputFormatter/OutputFormatter**: Transform data during import/export operations
- **ColumnInputFormatter/ColumnOutputFormatter**: Column-specific data transformations
- **CellReader**: Custom cell reading logic
- **Template Export**: Export data based on pre-designed Excel templates
- **Multi-Sheet Support**: Handle multiple sheets in a single workbook
- **Shadow Properties**: Define additional export columns not present in your models
- **Auto Column Width**: Automatic column width adjustment
- **Freeze Panes**: Set freeze panes for better data viewing
- **Filters**: Add auto-filters to your Excel sheets

### GetStarted

#### Installation

Install via NuGet Package Manager:

```bash
dotnet add package WeihanLi.Npoi
```

Or via Package Manager Console:

```powershell
Install-Package WeihanLi.Npoi
```

#### Quick Start

1. Export list/dataTable to Excel/csv

    ``` csharp
    var entities = new List<Entity>();
    
    // Export to Excel file
    entities.ToExcelFile(string excelPath);
    
    // Export to Excel bytes
    entities.ToExcelBytes(ExcelFormat excelFormat);
    
    // Export to CSV file
    entities.ToCsvFile(string csvPath);
    
    // Export to CSV bytes
    entities.ToCsvBytes();
    ```

2. Import Excel/csv to List

    ``` csharp
    // Read Excel first sheet content to List<T>
    var entityList = ExcelHelper.ToEntityList<T>(string excelPath);

    // Read Excel first sheet content to IEnumerable<T>
    var entityList = ExcelHelper.ToEntities<T>(string excelPath);

    // Read Excel specific sheet content to List<T>
    // You can customize header row index via sheet attribute or fluent api HasSheet
    var entityList1 = ExcelHelper.ToEntityList<T>(string excelPath, int sheetIndex);

    // Import CSV to List<T>
    var entityList2 = CsvHelper.ToEntityList<T>(string csvPath);
    var entityList3 = CsvHelper.ToEntityList<T>(byte[] csvBytes);
    ```

3. Import Excel/csv to DataTable

    ``` csharp
    // Read Excel to DataTable directly, by default read the first sheet content
    var dataTable = ExcelHelper.ToDataTable(string excelPath);

    // Read Excel workbook's specific sheet to DataTable
    var dataTableOfSheetIndex = ExcelHelper.ToDataTable(string excelPath, int sheetIndex);

    // Read Excel with custom header row index
    var dataTableOfSheetIndex = ExcelHelper.ToDataTable(string excelPath, int sheetIndex, int headerRowIndex);

    // Read Excel to DataTable using mapping relations and settings from typeof(T)
    var dataTableT = ExcelHelper.ToDataTable<T>(string excelPath);

    // Read CSV file data to DataTable
    var dataTable1 = CsvHelper.ToDataTable(string csvFilePath);
    ```

More Api documentation: <https://weihanli.github.io/WeihanLi.Npoi/api/WeihanLi.Npoi.html>

### Configuration

#### 1. Using Attributes

Add `ColumnAttribute` on the properties of your entity for export or import operations.

Add `SheetAttribute` on the entity to configure sheet settings. You can set the `StartRowIndex` as needed (default is `1`).

Example:

``` csharp
[Sheet(SheetName = "TestSheet", SheetIndex = 0, AutoColumnWidthEnabled = true)]
public class TestEntity
{
    [Column("ID", Index = 0)]
    public int PKID { get; set; }

    [Column("Bill Title", Index = 1)]
    public string BillTitle { get; set; }

    [Column("Bill Details", Index = 2)]
    public string BillDetails { get; set; }

    [Column("Created By", Index = 3)]
    public string CreatedBy { get; set; }

    [Column("Created Time", Index = 4, Formatter = "yyyy-MM-dd HH:mm:ss")]
    public DateTime CreatedTime { get; set; }
    
    [Column(IsIgnored = true)]
    public string InternalNote { get; set; }
}

public class TestEntity1
{
    [Column("Username")]
    public string Username { get; set; }

    [Column(IsIgnored = true)]
    public string PasswordHash { get; set; }

    [Column("Amount")]
    public decimal Amount { get; set; } = 1000M;

    [Column("WeChat OpenID")]
    public string WechatOpenId { get; set; }

    [Column("Is Active")]
    public bool IsActive { get; set; }
}
```

#### 2. Using FluentAPI (Recommended)

FluentAPI provides greater flexibility and more powerful configuration options.

Example:

``` csharp
var setting = FluentSettings.For<TestEntity>();

// Excel document settings
setting.HasAuthor("WeihanLi")
    .HasTitle("WeihanLi.Npoi test")
    .HasDescription("WeihanLi.Npoi test")
    .HasSubject("WeihanLi.Npoi test");

// Sheet configuration (sheetIndex, sheetName, startRowIndex, autoColumnWidth)
setting.HasSheetConfiguration(0, "SystemSettingsList", 1, true);

// Apply filters and freeze panes
// setting.HasFilter(0, 1).HasFreezePane(0, 1, 2, 1);

// Configure individual properties
setting.Property(_ => _.SettingId)
    .HasColumnIndex(0);

setting.Property(_ => _.SettingName)
    .HasColumnTitle("SettingName")
    .HasColumnIndex(1);

setting.Property(_ => _.DisplayName)
    .HasOutputFormatter((entity, displayName) => $"AAA_{entity.SettingName}_{displayName}")
    .HasInputFormatter((entity, originVal) => originVal.Split(new[] { '_' })[2])
    .HasColumnTitle("DisplayName")
    .HasColumnIndex(2);

setting.Property(_ => _.SettingValue)
    .HasColumnTitle("SettingValue")
    .HasColumnIndex(3);

setting.Property(_ => _.CreatedTime)
    .HasColumnTitle("CreatedTime")
    .HasColumnIndex(4)
    .HasColumnWidth(10)
    .HasColumnFormatter("yyyy-MM-dd HH:mm:ss");

setting.Property(_ => _.CreatedBy)
    .HasColumnInputFormatter(x => x += "_test")
    .HasColumnIndex(5)
    .HasColumnTitle("CreatedBy");

setting.Property(x => x.Enabled)
    .HasColumnInputFormatter(val => "Enabled".Equals(val))
    .HasColumnOutputFormatter(v => v ? "Enabled" : "Disabled");

// Shadow property - define a column that doesn't exist in the model
setting.Property("HiddenProp")
    .HasOutputFormatter((entity, val) => $"HiddenProp_{entity.PKID}");

// Ignore specific properties
setting.Property(_ => _.PKID).Ignored();
setting.Property(_ => _.UpdatedBy).Ignored();
setting.Property(_ => _.UpdatedTime).Ignored();
```

### Advanced Features

#### Template-based Export

Export data based on pre-designed Excel templates with placeholder support:

```csharp
entities.ToExcelFileByTemplate(
    templatePath: "path/to/template.xlsx",
    excelPath: "path/to/output.xlsx",
    extraData: new { Author = "WeihanLi", Title = "Export Result" }
);
```

Learn more: [Template Export Documentation](https://weihanli.github.io/WeihanLi.Npoi/articles/en/TemplateExport.html)

#### Multi-Sheet Export

Export multiple collections to different sheets in a single workbook:

```csharp
var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);
workbook.ImportData(collection1, sheetIndex: 0);
workbook.ImportData(collection2, sheetIndex: 1);
workbook.WriteToFile("multi-sheets.xlsx");
```

Learn more: [Multi-Sheet Documentation](https://weihanli.github.io/WeihanLi.Npoi/articles/en/MultiSheets.html)

#### Shadow Properties

Define additional export columns that don't exist in your model:

```csharp
var settings = FluentSettings.For<TestEntity>();
settings.Property("Employee ID")
    .HasOutputFormatter((entity, val) => $"{entity.UserFields[2].Value}");
settings.Property("Department")
    .HasOutputFormatter((entity, val) => $"{entity.UserFields[1].Value}");
```

Learn more: [Shadow Property Documentation](https://weihanli.github.io/WeihanLi.Npoi/articles/en/ShadowProperty.html)

### Documentation

- 📖 [Complete Documentation](https://weihanli.github.io/WeihanLi.Npoi/)
- 🚀 [Getting Started Guide](https://weihanli.github.io/WeihanLi.Npoi/articles/en/GetStarted.html)
- 📚 [API Reference](https://weihanli.github.io/WeihanLi.Npoi/api/WeihanLi.Npoi.html)
- 🌐 [Articles](https://weihanli.github.io/WeihanLi.Npoi/articles/intro.html)

### More

see some articles here: <https://weihanli.github.io/WeihanLi.Npoi/articles/intro.html>

more usage:

<details>
<summary>Get a workbook</summary>

``` csharp
// load excel workbook from file
var workbook = LoadExcel(string excelPath);

// prepare a workbook accounting to excelPath
var workbook = PrepareWorkbook(string excelPath);

// prepare a workbook accounting to excelPath and custom excel settings
var workbook = PrepareWorkbook(string excelPath, ExcelSetting excelSetting);

// prepare a workbook whether *.xls file
var workbook = PrepareWorkbook(bool isXls);

// prepare a workbook whether *.xls file and custom excel setting
var workbook = PrepareWorkbook(bool isXlsx, ExcelSetting excelSetting);
```

</details>

<details>
<summary>Rich extensions</summary>

``` csharp
List<TEntity> ToEntityList<TEntity>([NotNull]this IWorkbook workbook)

DataTable ToDataTable([NotNull]this IWorkbook workbook)

ISheet ImportData<TEntity>([NotNull] this ISheet sheet, DataTable dataTable)

int ImportData<TEntity>([NotNull] this IWorkbook workbook, IEnumerable<TEntity> list,
            int sheetIndex)

int ImportData<TEntity>([NotNull] this ISheet sheet, IEnumerable<TEntity> list)

int ImportData<TEntity>([NotNull] this IWorkbook workbook, [NotNull] DataTable dataTable,
            int sheetIndex)

ToExcelFile<TEntity>([NotNull] this IEnumerable<TEntity> entityList,
            [NotNull] string excelPath)

int ToExcelStream<TEntity>([NotNull] this IEnumerable<TEntity> entityList,
            [NotNull] Stream stream)

byte[] ToExcelBytes<TEntity>([NotNull] this IEnumerable<TEntity> entityList)

int ToExcelFile([NotNull] this DataTable dataTable, [NotNull] string excelPath)

int ToExcelStream([NotNull] this DataTable dataTable, [NotNull] Stream stream)

byte[] ToExcelBytes([NotNull] this DataTable dataTable)

byte[] ToExcelBytes([NotNull] this IWorkbook workbook)

int WriteToFile([NotNull] this IWorkbook workbook, string filePath)

object GetCellValue([NotNull] this ICell cell, Type propertyType)

T GetCellValue<T>([NotNull] this ICell cell)

void SetCellValue([NotNull] this ICell cell, object value)

byte[] ToCsvBytes<TEntity>(this IEnumerable<TEntity> entities, bool includeHeader)

ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath, bool includeHeader)

void ToCsvFile(this DataTable dt, string filePath, bool includeHeader)

byte[] ToCsvBytes(this DataTable dt, bool includeHeader)

```

</details>

### Samples

- [.NET Core Sample](https://github.com/WeihanLi/WeihanLi.Npoi/blob/dev/samples/DotNetCoreSample/Program.cs)
- [More Samples in Unit Tests](https://github.com/WeihanLi/WeihanLi.Npoi/blob/dev/test/WeihanLi.Npoi.Test/ExcelTest.cs)
- [Guide Posts](https://weihanli.github.io/WeihanLi.Npoi/articles/intro.html)

### Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Acknowledgements

- Thanks to all the [contributors](https://github.com/WeihanLi/WeihanLi.Npoi/graphs/contributors) and users of this project
- Thanks to [NPOI](https://github.com/tonyqus/npoi) for the excellent Excel library
- Thanks to [FluentExcel](https://github.com/Arch/FluentExcel/) for the FluentAPI inspiration
- Thanks to JetBrains for the free Rider license

### License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.

### Contact & Support

- 📧 Report issues: [GitHub Issues](https://github.com/WeihanLi/WeihanLi.Npoi/issues)
- 💬 Discussions: [GitHub Discussions](https://github.com/WeihanLi/WeihanLi.Npoi/discussions)
- 📦 NuGet Package: [WeihanLi.Npoi](https://www.nuget.org/packages/WeihanLi.Npoi/)
