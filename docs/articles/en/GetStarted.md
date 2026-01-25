# `WeihanLi.Npoi` Getting Started

## Introduction

`WeihanLi.Npoi` is an Excel import/export library based on NPOI that provides many useful extension methods and also supports CSV import/export.

- Import excel/csv data into `DataTable` or `List<TEntity>`
- Export `IEnumerable<TEntity>` or `DataTable` to Excel files, byte arrays, or streams
- Export `IEnumerable<TEntity>` or `DataTable` to CSV files or byte arrays
- Configuration through `Attributes` or `FluentAPI` (inspired by [FluentExcel](https://github.com/Arch/FluentExcel/))

## Basic Examples

``` csharp
internal class BaseModel
{
    public int Id { get; set; }
}

internal class Notice : BaseModel
{
    public string Title { get; set; }

    public string Content { get; set; }

    public DateTime PublishedAt { get; set; }

    public string Publisher { get; set; }
}
```

Basic import and export:

``` csharp
// entities excel import/export
[Theory]
[InlineData(ExcelFormat.Xls)]
[InlineData(ExcelFormat.Xlsx)]
public void BasicImportExportTest(ExcelFormat excelFormat)
{
    var list = new List<Notice>();
    for (var i = 0; i < 10; i++)
    {
        list.Add(new Notice()
        {
            Id = i + 1,
            Content = $"content_{i}",
            Title = $"title_{i}",
            PublishedAt = DateTime.UtcNow.AddDays(-i),
            Publisher = $"publisher_{i}"
        });
    }
    list.Add(new Notice() { Title = "nnnn" });
    list.Add(null);
    var excelBytes = list.ToExcelBytes(excelFormat);

    var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes, excelFormat);
    Assert.Equal(list.Count, importedList.Count);
    for (var i = 0; i < list.Count; i++)
    {
        if (list[i] == null)
        {
            Assert.Null(importedList[i]);
        }
        else
        {
            Assert.Equal(list[i].Id, importedList[i].Id);
            Assert.Equal(list[i].Title, importedList[i].Title);
            Assert.Equal(list[i].Content, importedList[i].Content);
            Assert.Equal(list[i].Publisher, importedList[i].Publisher);
            Assert.Equal(list[i].PublishedAt.ToTimeString(), importedList[i].PublishedAt.ToTimeString());
        }
    }
}

// DataTable Excel import/export
[Theory]
[InlineData(ExcelFormat.Xls)]
[InlineData(ExcelFormat.Xlsx)]
public void DataTableImportExportTest(ExcelFormat excelFormat)
{
    var dt = new DataTable();
    dt.Columns.AddRange(new[]
    {
        new DataColumn("Name"),
        new DataColumn("Age"),
        new DataColumn("Desc"),
    });
    for (var i = 0; i < 10; i++)
    {
        var row = dt.NewRow();
        row.ItemArray = new object[] { $"Test_{i}", i + 10, $"Desc_{i}" };
        dt.Rows.Add(row);
    }
    //
    var excelBytes = dt.ToExcelBytes(excelFormat);
    var importedData = ExcelHelper.ToDataTable(excelBytes, excelFormat);
    Assert.NotNull(importedData);
    Assert.Equal(dt.Rows.Count, importedData.Rows.Count);
    for (var i = 0; i < dt.Rows.Count; i++)
    {
        Assert.Equal(dt.Rows[i].ItemArray.Length, importedData.Rows[i].ItemArray.Length);
        for (var j = 0; j < dt.Rows[i].ItemArray.Length; j++)
        {
            Assert.Equal(dt.Rows[i].ItemArray[j], importedData.Rows[i].ItemArray[j]);
        }
    }
}

// entities csv import/export
[Fact]
public void BasicImportExportTest()
{
    var list = new List<Notice>();
    for (var i = 0; i < 10; i++)
    {
        list.Add(new Notice()
        {
            Id = i + 1,
            Content = $"content_{i}",
            Title = $"title_{i}",
            PublishedAt = DateTime.UtcNow.AddDays(-i),
            Publisher = $"publisher_{i}"
        });
    }
    list.Add(new Notice()
    {
        Id = 11,
        Content = $"content",
        Title = $"title",
        PublishedAt = DateTime.UtcNow.AddDays(1),
    });
    var csvBytes = list.ToCsvBytes();
    var importedList = CsvHelper.ToEntityList<Notice>(csvBytes);
    Assert.Equal(list.Count, importedList.Count);
    for (var i = 0; i < list.Count; i++)
    {
        Assert.Equal(list[i].Id, importedList[i].Id);
        Assert.Equal(list[i].Title ?? "", importedList[i].Title);
        Assert.Equal(list[i].Content ?? "", importedList[i].Content);
        Assert.Equal(list[i].Publisher ?? "", importedList[i].Publisher);
        Assert.Equal(list[i].PublishedAt.ToTimeString(), importedList[i].PublishedAt.ToTimeString());
    }
}
// DataTable csv import/export
[Fact]
public void DataTableImportExportTest()
{
    var dt = new DataTable();
    dt.Columns.AddRange(new[]
    {
        new DataColumn("Name"),
        new DataColumn("Age"),
        new DataColumn("Desc"),
    });
    for (var i = 0; i < 10; i++)
    {
        var row = dt.NewRow();
        row.ItemArray = new object[] { $"Test_{i}", i + 10, $"Desc_{i}" };
        dt.Rows.Add(row);
    }
    //
    var csvBytes = dt.ToCsvBytes();
    var importedData = CsvHelper.ToDataTable(csvBytes);
    Assert.NotNull(importedData);
    Assert.Equal(dt.Rows.Count, importedData.Rows.Count);
    for (var i = 0; i < dt.Rows.Count; i++)
    {
        Assert.Equal(dt.Rows[i].ItemArray.Length, importedData.Rows[i].ItemArray.Length);
        for (var j = 0; j < dt.Rows[i].ItemArray.Length; j++)
        {
            Assert.Equal(dt.Rows[i].ItemArray[j], importedData.Rows[i].ItemArray[j]);
        }
    }
}
```

## Custom Mapping and Configuration

Using Attributes:

``` csharp
internal class Model
{
    [Column("Hotel ID", Index = 0)]
    public string HotelId { get; set; }

    [Column("Order No", Index = 1)]
    public string OrderNo { get; set; }

    [Column("Hotel Name", Index = 2)]
    public string HotelName { get; set; }

    [Column("Customer Name", Index = 3)]
    public string CustomerName { get; set; }

    [Column(nameof(RoomType), Index = 4)]
    public string RoomType { get; set; }

    [Column(nameof(CheckInDate), Index = 5, Formatter = "yyyy/M/d")]
    public DateTime CheckInDate { get; set; }

    [Column(nameof(CheckOutDate), Index = 6, Formatter = "yyyy/M/d")]
    public DateTime CheckOutDate { get; set; }

    [Column(nameof(RoomNights), Index = 7)]
    public int RoomNights { get; set; }

    [Column(nameof(PaymentType), Index = 8)]
    public string PaymentType { get; set; }

    [Column(nameof(OrderAmount), Index = 9)]
    public decimal OrderAmount { get; set; }

    [Column(nameof(CommissionRate), Index = 10)]
    public decimal CommissionRate { get; set; }

    [Column(nameof(ServiceFee), Index = 11)]
    public decimal ServiceFee { get; set; }
}

[Sheet(SheetIndex = 0, SheetName = "TestSheet", AutoColumnWidthEnabled = true)]
internal class TestEntity2
{
    [Column(Index = 0)]
    public int Id { get; set; }

    [Column(Index = 1)]
    public string Title { get; set; }

    [Column(Index = 2, Width = 50)]
    public string Description { get; set; }

    [Column(Index = 3, Width = 20)]
    public string Extra { get; set; } = "{}";
}
```

Using FluentAPI (Recommended for greater flexibility):

``` csharp
var setting = FluentSettings.For<TestEntity>();
// ExcelSetting
setting.HasAuthor("WeihanLi")
    .HasTitle("WeihanLi.Npoi test")
    .HasDescription("WeihanLi.Npoi test")
    .HasSubject("WeihanLi.Npoi test");

setting.HasSheetConfiguration(0, "SystemSettingsList", 1, true); // sheet configuration

// setting
//     .HasFilter(0, 1) // Set filter on columns
//     .HasFreezePane(0, 1, 2, 1); // Set freeze pane

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
    .HasColumnWidth(10) // Set column width
    .HasColumnFormatter("yyyy-MM-dd HH:mm:ss");

setting.Property(_ => _.CreatedBy)
    .HasColumnInputFormatter(x => x += "_test")
    .HasColumnIndex(4)
    .HasColumnTitle("CreatedBy");

setting.Property(x => x.Enabled)
    .HasColumnInputFormatter(val => "Enabled".Equals(val))
    .HasColumnOutputFormatter(v => v ? "Enabled" : "Disabled");

setting.Property("ShadowProperty")
    .HasOutputFormatter((entity, val) => $"HiddenProp_{entity.PKID}");

setting.Property(_ => _.PKID).Ignored(); // ignore column
```
