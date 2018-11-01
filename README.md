# WeihanLi.Npoi [![WeihanLi.Npoi](https://img.shields.io/nuget/v/WeihanLi.Npoi.svg)](https://www.nuget.org/packages/WeihanLi.Npoi/)

## Build

[![Build Status](https://travis-ci.org/WeihanLi/WeihanLi.Npoi.svg?branch=master)](https://travis-ci.org/WeihanLi/WeihanLi.Npoi)

## Intro

NpoiExtensions for target framework net4.5 or netstandard2.0.

There' a lot of userful extensions for you, core fetures are as follows:

- mapping a excel file data to a `DataTable` or `List<TEntity>`
- export a `IEnumerable<TEntity>` or `DataTable` to Excel file or Excel file bytes or even write excel file stream to your stream
- export a `IEnumerable<TEntity>` or `DataTable` to csv file or bytes.

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

    it consider the first row of the sheet as the header, not for read,it will read data from next row.
    You can point out your header row through the exposed api if needed.

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

    void SetCellValue([NotNull] this ICell cell, object value)

    byte[] ToCsvBytes<TEntity>(this IEnumerable<TEntity> entities, bool includeHeader)

    ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath, bool includeHeader)

    void ToCsvFile(this DataTable dt, string filePath, bool includeHeader)

    byte[] ToCsvBytes(this DataTable dt, bool includeHeader)

    ```

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

    setting.HasSheetConfiguration(0, "System Settings");

    setting.HasFilter(0, 1)
        .HasFreezePane(0, 1, 2, 1);
    setting.Property(_ => _.SettingId)
        .HasColumnIndex(0);

    setting.Property(_ => _.SettingName)
        .HasColumnTitle("SettingName")
        .HasColumnIndex(1);

    setting.Property(_ => _.DisplayName)
        .HasColumnTitle("DisplayName")
        .HasColumnIndex(2);

    setting.Property(_ => _.SettingValue)
        .HasColumnTitle("SettingValue")
        .HasColumnIndex(3);

    setting.Property(_ => _.CreatedTime)
        .HasColumnTitle("CreatedTime")
        .HasColumnIndex(5)
        .HasColumnFormatter("yyyy-MM-dd HH:mm:ss");

    setting.Property(_ => _.CreatedBy)
        .HasColumnIndex(4)
        .HasColumnTitle("CreatedBy");

    setting.Property(_ => _.UpdatedBy).Ignored();
    setting.Property(_ => _.UpdatedTime).Ignored();
    setting.Property(_ => _.PKID).Ignored();
    ```

### Samples

- [dotnetcore sample](https://github.com/WeihanLi/WeihanLi.Npoi/blob/dev/samples/DotNetCoreSample/Program.cs)
- [dotnet sample](https://github.com/WeihanLi/WeihanLi.Npoi/blob/dev/samples/DotNetSample/Program.cs)

### Contact

Contact me: weihanli@outlook.com