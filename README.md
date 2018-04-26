# WeihanLi.Npoi [![WeihanLi.Npoi](https://img.shields.io/nuget/v/WeihanLi.Npoi.svg)](https://www.nuget.org/packages/WeihanLi.Npoi/)

## Build

[![Build Status](https://travis-ci.org/WeihanLi/WeihanLi.Npoi.svg?branch=master)](https://travis-ci.org/WeihanLi/WeihanLi.Npoi)

## Intro

Npoi 扩展,适用于.netframework4.5及以上和netstandard2.0, .netframework基于[NPOI](https://www.nuget.org/packages/NPOI/), .netstandard基于 [DotNetCore.NPOI](https://www.nuget.org/packages/DotNetCore.NPOI/)

NpoiExtensions for target framework net4.5 or netstandard2.0,for net45 basedon [NPOI](https://www.nuget.org/packages/NPOI/),for .netstandard basedon [DotNetCore.NPOI](https://www.nuget.org/packages/DotNetCore.NPOI/)

## Use

### Install

.NetFramework

``` bash
Install-Package WeihanLi.Npoi
```

.NetCore

``` bash
dotnet add package WeihanLi.Npoi
```

### GetStarted

1. LoadFromExcelFile

    it consider the first row of the sheet as the header not for read,it will read data from next row.You can point out your header row through the exposed api if needed.

    - Read Excel to DataSet

        ``` csharp
        // read excel to dataSet, read all sheets data to dataSet,by default it will read from the headerRowIndex(0) + 1
        var dataSet = ExcelHelper.ToDataSet(string excelPath);

        // read excel to dataSet, read all sheets data to dataSet,headerRowIndex is not for read,read from headerRowIndex+1
        var dataSet = ExcelHelper.ToDataSet(string excelPath, int headerRowIndex);
        ```

    - Read Excel to DataTable

        ``` csharp
        // read excel to dataTable directly,by default read the first sheet content
        var dataTable = ExcelHelper.ToDataTable(string excelPath);

        // read excel workbook's sheetIndex sheet to dataTable directly
        var dataTableOfSheetIndex = ExcelHelper.ToDataTable(string excelPath, int sheetIndex);

        // read excel workbook's sheetIndex sheet to dataTable,custom headerRowIndex
        var dataTableOfSheetIndex = ExcelHelper.ToDataTable(string excelPath, int sheetIndex, int headerRowIndex);

        // read excel to dataTable use mapping relations and settings from typeof(T),by default read the first sheet content
        var dataTableT = ExcelHelper.ToDataTable<T>(string excelPath);

        // ... sheetIndex and headerRowIndex is also supported like above
        ```

    - Read Excel to List

        ``` csharp
        // read excel first sheet content to a List<T>
        var entityList = ExcelHelper.ToEntityList<T>(string excelPath);

        // read excel sheetIndex sheet content to a List<T>
        // you can custom header row index via sheet attribute or fluent api HasSheet
        var entityList1 = ExcelHelper.ToEntityList<T>(string excelPath, int sheetIndex);
        ```

1. Get a workbook

``` csharp
// load excel workbook from file
var workbook = LoadExcel(string excelPath);

// prepare a workbook accounting to excelPath
var workbook = PrepareWorkbook(string excelPath);

// prepare a workbook accounting to excelPath and custom excel settings
var workbook = PrepareWorkbook(string excelPath, ExcelSetting excelSetting);

// prepare a workbook whether *.xlsx file
var workbook = PrepareWorkbook(bool isXlsx);

// prepare a workbook whether *.xlsx file and custom excel setting
var workbook = PrepareWorkbook(bool isXlsx, ExcelSetting excelSetting);
```

1. Rich extensions

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

SetCellValue([NotNull] this ICell cell, object value)

```

### Define Custom Mapping and settings

1. Attributes

    Add `ColumnAttribute` on the property of the entity which you used for export or import

    Add `SheetAttribute` on the entity which you used for export or import,you can set the `StartRowIndex` on your need(by default it is `1`)

    for example:

    ``` csharp
    public class TestEntity
    {
        [Column("PKID")]
        public int PKID { get; set; }

        [Column("账单标题")]
        public string BillTitle { get; set; }

        [Column("账单详情")]
        public string BillDetails { get; set; }

        [Column("创建人")]
        public string CreatedBy { get; set; }

        [Column("创建时间")]
        public DateTime CreatedTime { get; set; }
    }

    internal class TestEntity1
    {
        /// <summary>
        /// 用户名
        /// </summary>
        [Column("用户名")]
        public string Username { get; set; }

        [Column(IsIgnored = true)]
        public string PasswordHash { get; set; }

        [Column("可用余额")]
        public decimal Amount { get; set; } = 1000M;

        [Column("微信id")]
        public string WechatOpenId { get; set; }

        [Column("是否启用")]
        public bool IsActive { get; set; }
    }
    ```

1. FluentApi

    You can also use FluentApi above version `1.0.3`

    for example:

    ``` csharp
    var setting = ExcelHelper.SettingFor<TestEntity>();
    // ExcelSetting
    setting.HasAuthor("WeihanLi")
        .HasTitle("WeihanLi.Npoi test")
        .HasDescription("")
        .HasSubject("");

    setting.HasSheetConfiguration(0, "系统设置列表");

    setting.HasFilter(0, 1)
        .HasFreezePane(0, 1, 2, 1);
    setting.Property(_ => _.SettingId)
        .HasColumnIndex(0);

    setting.Property(_ => _.SettingName)
        .HasColumnTitle("设置名称")
        .HasColumnIndex(1);

    setting.Property(_ => _.DisplayName)
        .HasColumnTitle("设置显示名称")
        .HasColumnIndex(2);

    setting.Property(_ => _.SettingValue)
        .HasColumnTitle("设置值")
        .HasColumnIndex(3);

    setting.Property(_ => _.CreatedTime)
        .HasColumnTitle("创建时间")
        .HasColumnIndex(5)
        .HasColumnFormatter("yyyy-MM-dd HH:mm:ss");

    setting.Property(_ => _.CreatedBy)
        .HasColumnIndex(4)
        .HasColumnTitle("创建人");

    setting.Property(_ => _.UpdatedBy).Ignored();
    setting.Property(_ => _.UpdatedTime).Ignored();
    setting.Property(_ => _.PKID).Ignored();
    ```

### Samples

- [dotnetcore sample](https://github.com/WeihanLi/WeihanLi.Npoi/blob/dev/samples/DotNetCoreSample/Program.cs)
- [dotnet sample](https://github.com/WeihanLi/WeihanLi.Npoi/blob/dev/samples/DotNetSample/Program.cs)

### Contact

Contact me: weihanli@outlook.com