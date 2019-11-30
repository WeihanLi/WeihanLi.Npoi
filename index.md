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

