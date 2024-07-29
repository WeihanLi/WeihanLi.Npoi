# `WeihanLi.Npoi` 基础示例

## Intro

`WeihanLi.Npoi` 是基于 NPOI 扩展的 Excel 导入导出库，并提供了很多实用的扩展方法，也支持 CSV 的导入导出，

- 将 excel/csv 数据导入到 `DataTable` 或 `List<TEntity>`
- `IEnumerable<TEntity>` 或 `DataTable` 导出到 Excel，可以导出成 excel 文件或字节数组或者一个流
- `IEnumerable<TEntity>` 或 `DataTable` 导出到 csv 文件或者 csv 字节数组
- 通过 `Attribute` 或者 `FluentAPI`(借鉴了 [FluentExcel](https://github.com/Arch/FluentExcel/) 项目)

## BasicSample

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

基本的导入导出：

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

## 自定义映射关系，配置

使用 Attribute 配置：

``` csharp
internal class Model
{
    [Column("酒店编号", Index = 0)]
    public string HotelId { get; set; }

    [Column("订单号", Index = 1)]
    public string OrderNo { get; set; }

    [Column("酒店名称", Index = 2)]
    public string HotelName { get; set; }

    [Column("客户名称", Index = 3)]
    public string CustomerName { get; set; }

    [Column(nameof(房型名称), Index = 4)]
    public string 房型名称 { get; set; }

    [Column(nameof(入住日期), Index = 5, Formatter = "yyyy/M/d")]
    public DateTime 入住日期 { get; set; }

    [Column(nameof(离店日期), Index = 6, Formatter = "yyyy/M/d")]
    public DateTime 离店日期 { get; set; }

    [Column(nameof(间夜), Index = 7)]
    public int 间夜 { get; set; }

    [Column(nameof(支付类型), Index = 8)]
    public string 支付类型 { get; set; }

    [Column(nameof(订单金额), Index = 9)]
    public decimal 订单金额 { get; set; }

    [Column(nameof(佣金率), Index = 10)]
    public decimal 佣金率 { get; set; }

    [Column(nameof(服务费), Index = 11)]
    public decimal 服务费 { get; set; }
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

使用 FluentAPI 配置（推荐，更灵活）

``` csharp
var setting = FluentSettings.For<TestEntity>();
// ExcelSetting
setting.HasAuthor("WeihanLi")
    .HasTitle("WeihanLi.Npoi test")
    .HasDescription("WeihanLi.Npoi test")
    .HasSubject("WeihanLi.Npoi test");

setting.HasSheetConfiguration(0, "SystemSettingsList", 1, true); // sheet 配置

// setting
//     .HasFilter(0, 1) //在列上设置筛选
//     .HasFreezePane(0, 1, 2, 1); // 设置冻结区域

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
    .HasColumnWidth(10) // 设置列宽
    .HasColumnFormatter("yyyy-MM-dd HH:mm:ss");

setting.Property(_ => _.CreatedBy)
    .HasColumnInputFormatter(x => x += "_test")
    .HasColumnIndex(4)
    .HasColumnTitle("CreatedBy");

setting.Property(x => x.Enabled)
    .HasColumnInputFormatter(val => "启用".Equals(val))
    .HasColumnOutputFormatter(v => v ? "启用" : "禁用");

setting.Property("ShadowProperty")
    .HasOutputFormatter((entity, val) => $"HiddenProp_{entity.PKID}");

setting.Property(_ => _.PKID).Ignored(); // ignore column
```
