# WeihanLi.Npoi

[![WeihanLi.Npoi](https://img.shields.io/nuget/v/WeihanLi.Npoi)](https://www.nuget.org/packages/WeihanLi.Npoi/)

[![WeihanLi.Npoi Latest](https://img.shields.io/nuget/vpre/WeihanLi.Npoi)](https://www.nuget.org/packages/WeihanLi.Npoi/absoluteLatest)

## Build Status

[![Azure Pipeline Build Status](https://weihanli.visualstudio.com/Pipelines/_apis/build/status/WeihanLi.WeihanLi.Npoi?branchName=dev)](https://weihanli.visualstudio.com/Pipelines/_build/latest?definitionId=13&branchName=dev)

[![Github Build Status](https://github.com/WeihanLi/WeihanLi.Npoi/actions/workflows/dotnet.yml/badge.svg)](https://github.com/WeihanLi/WeihanLi.Npoi/actions/workflows/dotnet.yml)

## Intro

[NPOI](https://github.com/tonyqus/npoi) extensions based on target framework `netstandard2.0`.

There're a lot of useful extensions for you, core features are as follows:

- mapping a excel file data to a `DataTable` or `List<TEntity>`/`IEnumerable<T>`
- export `IEnumerable<TEntity>` or `DataTable` data to Excel file or Excel file bytes or even write excel file stream to your stream
- export `IEnumerable<TEntity>` or `DataTable` data to csv file or bytes.
- custom configuration/mappings by Attribute or FluentAPI(inspired by [FluentExcel](https://github.com/Arch/FluentExcel/))
- great flexibility by fluent InputFormatter/OutputFormatter/ColumnInputFormatter/ColumnOutputFormatter/CellReader

### GetStarted

1. Export list/dataTable to Excel/csv

    ``` csharp
    var entities = new List<Entity>();
    entities.ToExcelFile(string excelPath);
    entities.ToExcelBytes(ExcelFormat excelFormat);
    entities.ToCsvFile(string csvPath);
    entities.ToCsvBytes();
    ```

2. Read Excel/csv to List

    ``` csharp
    // read excel first sheet content to List<T>
    var entityList = ExcelHelper.ToEntityList<T>(string excelPath);

    // read excel first sheet content to IEnumerable<T>
    var entityList = ExcelHelper.ToEntities<T>(string excelPath);

    // read excel sheetIndex sheet content to a List<T>
    // you can custom header row index via sheet attribute or fluent api HasSheet
    var entityList1 = ExcelHelper.ToEntityList<T>(string excelPath, int sheetIndex);

    var entityList2 = CsvHelper.ToEntityList<T>(string csvPath);
    var entityList3 = CsvHelper.ToEntityList<T>(byte[] csvBytes);
    ```

3. Read Excel/csv to DataTable

    ``` csharp
    // read excel to dataTable directly,by default read the first sheet content
    var dataTable = ExcelHelper.ToDataTable(string excelPath);

    // read excel workbook's sheetIndex sheet to dataTable directly
    var dataTableOfSheetIndex = ExcelHelper.ToDataTable(string excelPath, int sheetIndex);

    // read excel workbook's sheetIndex sheet to dataTable,custom headerRowIndex
    var dataTableOfSheetIndex = ExcelHelper.ToDataTable(string excelPath, int sheetIndex, int headerRowIndex);

    // read excel to dataTable use mapping relations and settings from typeof(T),by default read the first sheet content
    var dataTableT = ExcelHelper.ToDataTable<T>(string excelPath);

    // read csv file data to dataTable
    var dataTable1 = CsvHelper.ToDataTable(string csvFilePath);
    ```

More Api here: <https://weihanli.github.io/WeihanLi.Npoi/api/WeihanLi.Npoi.html>

### Define Custom Mapping and settings

1. Attributes

    Add `ColumnAttribute` on the property of the entity which you used for export or import

    Add `SheetAttribute` on the entity which you used for export or import,you can set the `StartRowIndex` on your need(by default it is `1`)

    for example:

    ``` csharp
    public class TestEntity
    {
        [Column("Id")]
        public int PKID { get; set; }

        [Column("Bill Title")]
        public string BillTitle { get; set; }

        [Column("Bill Details")]
        public string BillDetails { get; set; }

        [Column("CreatedBy")]
        public string CreatedBy { get; set; }

        [Column("CreatedTime")]
        public DateTime CreatedTime { get; set; }
    }

    public class TestEntity1
    {
        [Column("Username")]
        public string Username { get; set; }

        [Column(IsIgnored = true)]
        public string PasswordHash { get; set; }

        [Column("Amount")]
        public decimal Amount { get; set; } = 1000M;

        [Column("WechatOpenId")]
        public string WechatOpenId { get; set; }

        [Column("IsActive")]
        public bool IsActive { get; set; }
    }
    ```

1. FluentApi (Recommended)

    You can use FluentApi for great flexibility

    for example:

    ``` csharp
    var setting = FluentSettings.For<TestEntity>();
    // ExcelSetting
    setting.HasAuthor("WeihanLi")
        .HasTitle("WeihanLi.Npoi test")
        .HasDescription("WeihanLi.Npoi test")
        .HasSubject("WeihanLi.Npoi test");

    setting.HasSheetConfiguration(0, "SystemSettingsList", true);
    // setting.HasSheetConfiguration(1, "SystemSettingsList", 1, true);

    // setting.HasFilter(0, 1).HasFreezePane(0, 1, 2, 1);

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
        .HasColumnIndex(4)
        .HasColumnTitle("CreatedBy");

    setting.Property(x => x.Enabled)
        .HasColumnInputFormatter(val => "Enabled".Equals(val))
        .HasColumnOutputFormatter(v => v ? "Enabled" : "Disabled");

    setting.Property("HiddenProp")
        .HasOutputFormatter((entity, val) => $"HiddenProp_{entity.PKID}");

    setting.Property(_ => _.PKID).Ignored();
    setting.Property(_ => _.UpdatedBy).Ignored();
    setting.Property(_ => _.UpdatedTime).Ignored();
    ```

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

- [dotnetcore sample](https://github.com/WeihanLi/WeihanLi.Npoi/blob/dev/samples/DotNetCoreSample/Program.cs)
- [More samples in unit test](https://github.com/WeihanLi/WeihanLi.Npoi/blob/dev/test/WeihanLi.Npoi.Test/ExcelTest.cs)
- [Guide posts](https://weihanli.github.io/WeihanLi.Npoi/articles/intro.html)

### Acknowledgements

- Thanks for the contributors and users for this project
- Thanks JetBrains for the free Rider license
