# WeihanLi.Npoi

## Intro

[NPOI](https://github.com/tonyqus/npoi) extensions for target framework net45 and netstandard2.0, mainly for import and export

There' a lot of userful extensions for you, core fetures are as follows:

- mapping a excel file data to a `DataTable` or `List<TEntity>`
- export `IEnumerable<TEntity>` or `DataTable` data to Excel file or Excel file bytes or even write excel file stream to your stream
- export `IEnumerable<TEntity>` or `DataTable` data to csv file or bytes.
- custom configuration/mappings by Attribute or FluentAPI(inspired by [FluentExcel](https://github.com/Arch/FluentExcel/))


## Sample

``` csharp
var list2 = new List<TestEntity2>();
list2.Add(null);
for (var i = 0; i < 10; i++)
{
    list2.Add(new TestEntity2
    {
        Id = i + 1,
        Title = $"Title_{i}",
        Description = $"{Enumerable.Range(1, 200).StringJoin(",")}__{i}",
    });
}
list2.Add(new TestEntity2()
{
    Id = 999,
    Title = $"{Enumerable.Repeat(1, 10).StringJoin(",")}",
    Description = null
});
var tempDirPath = $@"{Environment.GetEnvironmentVariable("USERPROFILE")}\Desktop\temp\test";
list2.ToExcelFile($@"{tempDirPath}\testEntity2.xlsx");

var listTemp = ExcelHelper.ToEntityList<TestEntity2>($@"{tempDirPath}\testEntity2.xlsx");
var dataTableTemp = ExcelHelper.ToDataTable($@"{tempDirPath}\testEntity2.xlsx");

Console.WriteLine("Press Enter to continue...");
Console.ReadLine();

var entities = new List<TestEntity>()
{
    new TestEntity()
    {
        PKID = 1,
        SettingId = Guid.NewGuid(),
        SettingName = "Setting1",
        SettingValue = "Value1",
        DisplayName = "ddd1"
    },
    new TestEntity()
    {
        PKID = 2,
        SettingId = Guid.NewGuid(),
        SettingName = "Setting2",
        SettingValue = "Value2"
    },
};
var csvFilePath = $@"{tempDirPath}\test.csv";
entities.ToExcelFile(csvFilePath.Replace(".csv", ".xlsx"));
var entitiesT0 = ExcelHelper.ToEntityList<TestEntity>(csvFilePath.Replace(".csv", ".xlsx"));
entities.ToCsvFile(csvFilePath);

var dataTable = entities.ToDataTable();
dataTable.ToCsvFile(csvFilePath.Replace(".csv", ".datatable.csv"));
var dt = CsvHelper.ToDataTable(csvFilePath.Replace(".csv", ".datatable.csv"));
Console.WriteLine(dt.Columns.Count);
var entities1 = CsvHelper.ToEntityList<TestEntity>(csvFilePath);
entities1[1].DisplayName = ",tadadada";
entities1[0].SettingValue = "value2,345";
entities1.ToCsvFile(csvFilePath.Replace(".csv", ".1.csv"));
entities1.ToDataTable().ToCsvFile(csvFilePath.Replace(".csv", ".1.datatable.csv"));

var list = CsvHelper.ToEntityList<TestEntity>(csvFilePath.Replace(".csv", ".1.csv"));
dt = CsvHelper.ToDataTable(csvFilePath.Replace(".csv", ".1.datatable.csv"));
Console.WriteLine(dt.Columns.Count);
var entities2 = CsvHelper.ToEntityList<TestEntity>(csvFilePath.Replace(".csv", ".1.csv"));

entities.ToExcelFile(csvFilePath.Replace(".csv", ".xlsx"));

var vals = new[] { 1, 2, 3, 5, 4 };
vals.ToCsvFile(csvFilePath);

var numList = CsvHelper.ToEntityList<int>(csvFilePath);
Console.WriteLine(numList.StringJoin(","));

```

## Configuration

1. Attributes

    Add `ColumnAttribute` on the property of the entity which you used for export or import

    Add `SheetAttribute` on the entity which you used for export or import,you can set the `StartRowIndex` on your need(by default it is `1`)

    for example:

    ``` csharp
    public class TestEntity
    {
        [Column("Id")]
        public int PKID { get; set; }

        [Column("BillTitle")]
        public string BillTitle { get; set; }

        [Column("BillDetails")]
        public string BillDetails { get; set; }

        [Column("CreatedByUser")]
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

1. FluentApi (Recommend)

    You can use FluentApi also

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
        .HasOutputFormatter((entity, displayName) => $"AAA_{entity.SettingName}_{displayName}")
        .HasInputFormatter((entity, originVal) => originVal.Split(new[] { '_' })[2])
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
