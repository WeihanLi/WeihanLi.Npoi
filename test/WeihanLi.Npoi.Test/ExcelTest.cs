// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;
using WeihanLi.Common;
using WeihanLi.Common.Models;
using WeihanLi.Common.Services;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Attributes;
using WeihanLi.Npoi.Test.Models;
using Xunit;

namespace WeihanLi.Npoi.Test;

public class ExcelTest
{
    [Theory]
    [ExcelFormatData]
    public void BasicImportExportTest(ExcelFormat excelFormat)
    {
        var list = new List<Notice?>();
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
        var noticeSetting = FluentSettings.For<Notice>();
        lock (noticeSetting)
        {
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
                    Assert.NotNull(importedList[i]);
                    var sourceItem = list[i]!;
                    var item = importedList[i]!;
                    Assert.Equal(sourceItem.Id, item.Id);
                    Assert.Equal(sourceItem.Title, item.Title);
                    Assert.Equal(sourceItem.Content, item.Content);
                    Assert.Equal(sourceItem.Publisher, item.Publisher);
                    Assert.Equal(sourceItem.PublishedAt.ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());
                }
            }
        }
    }

    [Theory]
    [ExcelFormatData]
    public void BasicImportExportTestWithEmptyValue(ExcelFormat excelFormat)
    {
        var list = new List<Notice?>();
        for (var i = 0; i < 10; i++)
        {
            list.Add(new Notice()
            {
                Id = i + 1,
                Content = i < 3 ? $"content_{i}" : string.Empty,
                Title = $"title_{i}",
                PublishedAt = DateTime.UtcNow.AddDays(-i),
                Publisher = i < 3 ? $"publisher_{i}" : null
            });
        }
        list.Add(new Notice() { Title = "nnnn" });
        list.Add(null);
        var noticeSetting = FluentSettings.For<Notice>();
        lock (noticeSetting)
        {
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
                    Assert.NotNull(importedList[i]);
                    var sourceItem = list[i]!;
                    var item = importedList[i]!;
                    Assert.Equal(sourceItem.Id, item.Id);
                    Assert.Equal(sourceItem.Title, item.Title);
                    Assert.Equal(sourceItem.Content, item.Content);
                    Assert.Equal(sourceItem.Publisher, item.Publisher);
                    Assert.Equal(sourceItem.PublishedAt.ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());
                }
            }
        }
    }

    [Theory]
    [ExcelFormatData]
    public void BasicImportExportWithoutHeaderTest(ExcelFormat excelFormat)
    {
        var list = new List<Notice?>();
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

        var noticeSetting = FluentSettings.For<Notice>();
        lock (noticeSetting)
        {
            noticeSetting.HasSheetConfiguration(0, "test", 0);

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
                    Assert.NotNull(importedList[i]);
                    var sourceItem = list[i]!;
                    var item = importedList[i]!;
                    Assert.Equal(sourceItem.Id, item.Id);
                    Assert.Equal(sourceItem.Title, item.Title);
                    Assert.Equal(sourceItem.Content, item.Content);
                    Assert.Equal(sourceItem.Publisher, item.Publisher);
                    Assert.Equal(sourceItem.PublishedAt.ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());
                }
            }

            noticeSetting.HasSheetConfiguration(0, "test", 1);
        }
    }

    [Theory]
    [ExcelFormatData]
    public void ImportWithNotSpecificColumnIndex(ExcelFormat excelFormat)
    {
        IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
        {
            Id = i + 1,
            Content = $"content_{i}",
            Title = $"title_{i}",
            PublishedAt = DateTime.UtcNow.AddDays(-i),
            Publisher = $"publisher_{i}"
        }).ToArray();
        //
        var noticeSetting = FluentSettings.For<Notice>();
        lock (noticeSetting)
        {
            var excelBytes = list.ToExcelBytes(excelFormat);

            noticeSetting.Property(_ => _.Publisher)
                .HasColumnIndex(4);
            noticeSetting.Property(_ => _.PublishedAt)
                .HasColumnIndex(3);

            var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes, excelFormat);
            Assert.Equal(list.Count, importedList.Count);
            for (var i = 0; i < list.Count; i++)
            {
                if (importedList[i] == null)
                {
                    Assert.Null(list[i]);
                }
                else
                {
                    var item = importedList[i];
                    Assert.NotNull(item);
                    Guard.NotNull(item);
                    var sourceItem = list[i];
                    Assert.Equal(sourceItem.Id, item.Id);
                    Assert.Equal(sourceItem.Title, item.Title);
                    Assert.Equal(sourceItem.Content, item.Content);
                    Assert.Equal(sourceItem.Publisher, item.Publisher);
                    Assert.Equal(sourceItem.PublishedAt.ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());
                }
            }

            noticeSetting.Property(_ => _.Publisher)
                .HasColumnIndex(3);
            noticeSetting.Property(_ => _.PublishedAt)
                .HasColumnIndex(4);
        }
    }

    [Theory]
    [ExcelFormatData]
    public void ShadowPropertyTest(ExcelFormat excelFormat)
    {
        IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
        {
            Id = i + 1,
            Content = $"content_{i}",
            Title = $"title_{i}",
            PublishedAt = DateTime.UtcNow.AddDays(-i),
            Publisher = $"publisher_{i}"
        }).ToArray();

        var noticeSetting = FluentSettings.For<Notice>();
        lock (noticeSetting)
        {
            noticeSetting.Property<string>("ShadowProperty")
                .HasOutputFormatter((x, _) => $"{x?.Id}...")
                ;

            var excelBytes = list.ToExcelBytes(excelFormat);
            // list.ToExcelFile($"{Directory.GetCurrentDirectory()}/output.xlsx");

            var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes, excelFormat);
            Assert.Equal(list.Count, importedList.Count);
            for (var i = 0; i < list.Count; i++)
            {
                var item = importedList[i];
                Assert.NotNull(item);
                Guard.NotNull(item);
                var sourceItem = list[i];
                Assert.Equal(sourceItem.Id, item.Id);
                Assert.Equal(sourceItem.Title, item.Title);
                Assert.Equal(sourceItem.Content, item.Content);
                Assert.Equal(sourceItem.Publisher, item.Publisher);
                Assert.Equal(sourceItem.PublishedAt.ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());
            }
        }
    }

    [Theory]
    [ExcelFormatData]
    public void IgnoreInheritPropertyTest(ExcelFormat excelFormat)
    {
        IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
        {
            Id = i + 1,
            Content = $"content_{i}",
            Title = $"title_{i}",
            PublishedAt = DateTime.UtcNow.AddDays(-i),
            Publisher = $"publisher_{i}"
        }).ToArray();

        var settings = FluentSettings.For<Notice>();
        lock (settings)
        {
            settings.Property(x => x.Id).Ignored();

            var excelBytes = list.ToExcelBytes(excelFormat);
            // list.ToExcelFile($"{Directory.GetCurrentDirectory()}/ttt.xls");
            var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes, excelFormat);
            Assert.Equal(list.Count, importedList.Count);
            for (var i = 0; i < list.Count; i++)
            {
                if (importedList[i] == null)
                {
                    Assert.Null(list[i]);
                }
                else
                {
                    var item = importedList[i];
                    Assert.NotNull(item);
                    Guard.NotNull(item);
                    var sourceItem = list[i];
                    //Assert.Equal(sourceItem.Id, item.Id);
                    Assert.Equal(sourceItem.Title, item.Title);
                    Assert.Equal(sourceItem.Content, item.Content);
                    Assert.Equal(sourceItem.Publisher, item.Publisher);
                    Assert.Equal(sourceItem.PublishedAt.ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());
                }
            }

            settings.Property(_ => _.Id)
                .Ignored(false)
                .HasColumnIndex(0);
        }
    }

    [Theory]
    [ExcelFormatData]
    public void ColumnInputFormatterTest(ExcelFormat excelFormat)
    {
        IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
        {
            Id = i + 1,
            Content = $"content_{i}",
            Title = $"title_{i}",
            PublishedAt = DateTime.UtcNow.AddDays(-i),
            Publisher = $"publisher_{i}"
        }).ToArray();

        var excelBytes = list.ToExcelBytes(excelFormat);

        var settings = FluentSettings.For<Notice>();
        lock (settings)
        {
            settings.Property(x => x.Title).HasColumnInputFormatter(x => $"{x}_Test");

            var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes, excelFormat);
            Assert.Equal(list.Count, importedList.Count);
            for (var i = 0; i < list.Count; i++)
            {
                var item = importedList[i];
                Assert.NotNull(item);
                Guard.NotNull(item);

                Assert.Equal(list[i].Id, item.Id);
                Assert.Equal(list[i].Title + "_Test", item.Title);
                Assert.Equal(list[i].Content, item.Content);
                Assert.Equal(list[i].Publisher, item.Publisher);
                Assert.Equal(list[i].PublishedAt.ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());
            }

            settings.Property(_ => _.Title).HasColumnInputFormatter(null);
        }
    }

    [Theory]
    [ExcelFormatData]
    public void InputOutputColumnFormatterTest(ExcelFormat excelFormat)
    {
        IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
        {
            Id = i + 1,
            Content = $"content_{i}",
            Title = $"title_{i}",
            PublishedAt = DateTime.UtcNow.AddDays(-i),
            Publisher = $"publisher_{i}"
        }).ToArray();

        var settings = FluentSettings.For<Notice>();
        lock (settings)
        {
            settings.Property(x => x.Id)
                .HasColumnOutputFormatter(x => $"{x}_Test")
                .HasColumnInputFormatter(x => Convert.ToInt32(x?.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries)[0]));
            var excelBytes = list.ToExcelBytes(excelFormat);

            var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes, excelFormat);
            Assert.Equal(list.Count, importedList.Count);
            for (var i = 0; i < list.Count; i++)
            {
                var item = importedList[i];
                Assert.NotNull(item);
                Guard.NotNull(item);
                var sourceItem = list[i];
                Assert.Equal(sourceItem.Id, item.Id);
                Assert.Equal(sourceItem.Title, item.Title);
                Assert.Equal(sourceItem.Content, item.Content);
                Assert.Equal(sourceItem.Publisher, item.Publisher);
                Assert.Equal(sourceItem.PublishedAt.ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());
            }

            settings.Property(x => x.Id)
                .HasColumnOutputFormatter(null)
                .HasColumnInputFormatter(null);
        }
    }

    [Theory]
    [ExcelFormatData]
    public void DataValidationTest(ExcelFormat excelFormat)
    {
        IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
        {
            Id = i + 1,
            Content = $"content_{i}",
            Title = $"title_{i}",
            PublishedAt = DateTime.UtcNow.AddDays(-i),
            Publisher = $"publisher_{i}"
        }).ToArray();
        var excelBytes = list.ToExcelBytes(excelFormat);

        var settings = FluentSettings.For<Notice>();
        lock (settings)
        {
            settings.WithDataFilter(x => x?.Id > 5);

            var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes, excelFormat);
            Assert.Equal(list.Count(x => x.Id > 5), importedList.Count);

            int i = 0, k = 0;
            while (list[k].Id != importedList[i]?.Id)
            {
                k++;
            }

            for (; i < importedList.Count; i++, k++)
            {
                var item = importedList[i];
                Assert.NotNull(item);
                Guard.NotNull(item);
                var sourceItem = list[k];
                Assert.Equal(sourceItem.Id, item.Id);
                Assert.Equal(sourceItem.Title, item.Title);
                Assert.Equal(sourceItem.Content, item.Content);
                Assert.Equal(sourceItem.Publisher, item.Publisher);
                Assert.Equal(sourceItem.PublishedAt.ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());

            }

            settings.WithDataFilter(null);
        }
    }

    [Theory]
    [ExcelFormatData]
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

    [Theory]
    [ExcelFormatData]
    public void DataTableImportExportWithEmptyValueTest(ExcelFormat excelFormat)
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
            row.ItemArray = new object?[]
            {
                    i < 4 ? $"Test_{i}" : null,
                    i + 10,
                    i % 2 == 0 ? $"Desc_{i}" : string.Empty
            };
            dt.Rows.Add(row);
        }
        //
        var excelBytes = dt.ToExcelBytes(excelFormat);
        var importedData = ExcelHelper.ToDataTable(excelBytes, excelFormat);
        Assert.NotNull(importedData);
        Assert.Equal(dt.Rows.Count, importedData.Rows.Count);
        Assert.Equal(dt.Columns.Count, importedData.Columns.Count);
        for (var i = 0; i < dt.Rows.Count; i++)
        {
            Assert.Equal(dt.Rows[i].ItemArray.Length, importedData.Rows[i].ItemArray.Length);
            for (var j = 0; j < dt.Columns.Count; j++)
            {
                Assert.Equal(dt.Rows[i].ItemArray[j], importedData.Rows[i].ItemArray[j]);
            }
        }
    }

    [Theory]
    [ExcelFormatData]
    public void ExcelImportWithFormula(ExcelFormat excelFormat)
    {
        var setting = FluentSettings.For<ExcelFormulaTestModel>();
        setting.HasSheetConfiguration(0, "Test", 0);
        setting.Property(x => x.Num1).HasColumnIndex(0);
        setting.Property(x => x.Num2).HasColumnIndex(1);
        setting.Property(x => x.Sum).HasColumnIndex(2);

        var workbook = ExcelHelper.PrepareWorkbook(excelFormat);
        var sheet = workbook.CreateSheet();
        var row = sheet.CreateRow(0);
        row.CreateCell(0, CellType.Numeric).SetCellValue(1);
        row.CreateCell(1, CellType.Numeric).SetCellValue(2);
        row.CreateCell(2, CellType.Formula).SetCellFormula("$A1+$B1");
        var excelBytes = workbook.ToExcelBytes();
        var list = ExcelHelper.ToEntityList<ExcelFormulaTestModel>(excelBytes, excelFormat);
        Assert.NotNull(list);
        Assert.NotEmpty(list);
        Assert.NotNull(list[0]);
        Assert.Equal(1, list[0]!.Num1);
        Assert.Equal(2, list[0]!.Num2);
        Assert.Equal(3, list[0]!.Sum);
    }

    // ReSharper disable UnusedAutoPropertyAccessor.Local
    private class ExcelFormulaTestModel
    {
        public int Num1 { get; set; }
        public int Num2 { get; set; }

        public int Sum { get; set; }
    }

    [Theory]
    [ExcelFormatData]
    public void ExcelImportWithCellFilter(ExcelFormat excelFormat)
    {
        IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
        {
            Id = i + 1,
            Content = $"content_{i}",
            Title = $"title_{i}",
            PublishedAt = DateTime.UtcNow.AddDays(-i),
            Publisher = $"publisher_{i}"
        }).ToArray();
        var excelBytes = list.ToExcelBytes(excelFormat);

        var settings = FluentSettings.For<Notice>();
        lock (settings)
        {
            settings.HasSheetSetting(setting =>
            {
                setting.CellFilter = cell => cell.ColumnIndex == 0;
            });

            var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes, excelFormat);
            Assert.Equal(list.Count, importedList.Count);
            for (var i = 0; i < list.Count; i++)
            {
                var item = importedList[i];
                Assert.NotNull(item);
                Guard.NotNull(item);

                Assert.Equal(list[i].Id, item.Id);
                Assert.Null(item.Title);
                Assert.Null(item.Content);
                Assert.Null(item.Publisher);
                Assert.Equal(default(DateTime).ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());
            }

            settings.HasSheetSetting(setting =>
            {
                setting.CellFilter = _ => true;
            });
        }
    }

    [Theory]
    [ExcelFormatData]
    public void ExcelImportWithCellFilterAttributeTest(ExcelFormat excelFormat)
    {
        IReadOnlyList<CellFilterAttributeTest> list = Enumerable.Range(0, 10).Select(i => new CellFilterAttributeTest()
        {
            Id = i + 1,
            Description = $"content_{i}",
            Name = $"title_{i}",
        }).ToArray();
        var excelBytes = list.ToExcelBytes(excelFormat);
        var importedList = ExcelHelper.ToEntityList<CellFilterAttributeTest>(excelBytes, excelFormat);
        Assert.NotNull(importedList);
        Assert.Equal(list.Count, importedList.Count);
        for (var i = 0; i < importedList.Count; i++)
        {
            Assert.NotNull(importedList[i]);
            var item = importedList[i]!;
            Assert.Equal(list[i].Id, item.Id);
            Assert.Equal(list[i].Name, item.Name);
            Assert.Null(item.Description);
        }
    }

    [Sheet(SheetName = "test", AutoColumnWidthEnabled = true, StartColumnIndex = 0, EndColumnIndex = 1)]
    private class CellFilterAttributeTest
    {
        [Column(Index = 0)]
        public int Id { get; set; }

        [Column(Index = 1)]
        public string? Name { get; set; }

        [Column(Index = 2)]
        public string? Description { get; set; }
    }

    [Theory]
    [InlineData(ExcelFormat.Xls, 1000, 1)]
    [InlineData(ExcelFormat.Xls, 65536, 2)]
    [InlineData(ExcelFormat.Xls, 132_000, 3)]
    //[InlineData(ExcelFormat.Xls, 1_000_000, 16)]
    //[InlineData(ExcelFormat.Xlsx, 1_048_576, 2)]
    public void EntityListAutoSplitSheetsTest(ExcelFormat excelFormat, int rowsCount, int expectedSheetCount)
    {
        var list = Enumerable.Range(1, rowsCount)
            .Select(x => new Notice()
            {
                Id = x,
                Content = $"content_{x}",
                Title = $"title_{x}",
                Publisher = $"publisher_{x}"
            })
            .ToArray();

        var bytes = list.ToExcelBytes(excelFormat);
        var workbook = ExcelHelper.LoadExcel(bytes, excelFormat);
        Assert.Equal(expectedSheetCount, workbook.NumberOfSheets);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls, 1000, 1)]
    [InlineData(ExcelFormat.Xls, 65536, 2)]
    [InlineData(ExcelFormat.Xls, 132_000, 3)]
    //[InlineData(ExcelFormat.Xls, 1_000_000, 16)]
    //[InlineData(ExcelFormat.Xlsx, 1_048_576, 2)]
    public void DataTableAutoSplitSheetsTest(ExcelFormat excelFormat, int rowsCount, int expectedSheetCount)
    {
        var dataTable = new DataTable();
        dataTable.Columns.Add(new DataColumn("Id", typeof(int)));
        for (var i = 0; i < rowsCount; i++)
        {
            var row = dataTable.NewRow();
            row.ItemArray = new object[]
            {
                    i+1
            };
            dataTable.Rows.Add(row);
        }
        Assert.Equal(rowsCount, dataTable.Rows.Count);
        var bytes = dataTable.ToExcelBytes(excelFormat);
        var workbook = ExcelHelper.LoadExcel(bytes, excelFormat);
        Assert.Equal(expectedSheetCount, workbook.NumberOfSheets);
    }

    [Theory]
    [InlineData(@"TestData/EmptyColumns/emptyColumns.xls", ExcelFormat.Xls)]
    [InlineData(@"TestData/EmptyColumns/emptyColumns.xlsx", ExcelFormat.Xlsx)]
    public void DataTableImportExportTestWithFirstColumnsEmpty(string file, ExcelFormat excelFormat)
    {
        // Arrange
        var excelBytes = File.ReadAllBytes(file);

        // Act
        var importedData = ExcelHelper.ToDataTable(excelBytes, excelFormat);

        // Assert
        var dt = new DataTable();
        dt.Columns.AddRange(new[]
        {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C"),
                new DataColumn("D"),
            });

        dt.AddNewRow(new object[] { "", "", "3", "4" });
        dt.AddNewRow(new object[] { "", "2", "3", "" });
        dt.AddNewRow(new object[] { "1", "2", "", "" });
        dt.AddNewRow(new object[] { "1", "2", "3", "4" });

        Assert.NotNull(importedData);

        Assert.Equal(4, importedData.Rows.Count);

        importedData.AssertEquals(dt);
    }

    [Theory]
    [InlineData(@"TestData/NonStringColumns/nonStringColumns.xls", ExcelFormat.Xls)]
    [InlineData(@"TestData/NonStringColumns/nonStringColumns.xlsx", ExcelFormat.Xlsx)]
    public void DataTableImportExportTestWithNonStringColumns(string file, ExcelFormat excelFormat)
    {
        // Arrange
        var excelBytes = File.ReadAllBytes(file);

        // Act
        var importedData = ExcelHelper.ToDataTable(excelBytes, excelFormat);

        // Assert
        var dt = new DataTable();
        dt.Columns.AddRange(new[]
        {
                new DataColumn("A"),
                new DataColumn("1000"),
                new DataColumn("TRUE"), // Excel value will loaded as "True".
                new DataColumn(DateTime.ParseExact("15/08/2021", "dd/MM/yyyy", CultureInfo.InvariantCulture).ToShortDateString()),
            });

        dt.AddNewRow(new object[] { "1", "2", "3", "4" });

        Assert.NotNull(importedData);

        Assert.Equal(1, importedData.Rows.Count);

        importedData.AssertEquals(dt);
    }

    [Theory]
    [InlineData(@"TestData/EmptyRows/emptyRows.xls", ExcelFormat.Xls)]
    [InlineData(@"TestData/EmptyRows/emptyRows.xlsx", ExcelFormat.Xlsx)]
    public void DataTableImportExportTestWithoutEmptyRowsAndAdditionalColumns(string file, ExcelFormat excelFormat)
    {
        // Arrange
        var excelBytes = File.ReadAllBytes(file);

        // Act
        var importedData = ExcelHelper.ToDataTable(excelBytes, excelFormat, removeEmptyRows: true, maxColumns: 3);

        // Assert
        var dt = new DataTable();
        dt.Columns.AddRange(new[]
        {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C"),
            });

        dt.AddNewRow(new object[] { "1", "2", "3" });
        dt.AddNewRow(new object[] { "1", "", "" });
        dt.AddNewRow(new object[] { "1", "2", "3" });
        dt.AddNewRow(new object[] { "", "2", "3" });

        Assert.NotNull(importedData);

        Assert.Equal(4, importedData.Rows.Count);

        importedData.AssertEquals(dt);
    }

    [Theory]
    [ExcelFormatData]
    public async Task ImageImportExportTest(ExcelFormat excelFormat)
    {
        using var httpClient = new HttpClient();
        var imageBytes = await httpClient.GetByteArrayAsync("https://www.nuget.org/profiles/weihanli/avatar?imageSize=64");
        var list = Enumerable.Range(1, 5)
            .Select(x => new ImageTest() { Id = x, Image = imageBytes })
            .ToList();
        var excelBytes = list.ToExcelBytes(excelFormat);
        var importResult = ExcelHelper.ToEntityList<ImageTest>(excelBytes, excelFormat);
        Assert.NotNull(importResult);
        Assert.Equal(list.Count, importResult.Count);
        for (var i = 0; i < list.Count; i++)
        {
            Assert.NotNull(importResult[i]);
            var result = importResult[i]!;
            Assert.Equal(list[i].Id, result.Id);
            Assert.True(list[i].Image.SequenceEqual(result.Image));
        }
    }

    [Theory]
    [ExcelFormatData]
    public async Task ImageImportExportPictureDataTest(ExcelFormat excelFormat)
    {
        using var httpClient = new HttpClient();
        var imageBytes = await httpClient.GetByteArrayAsync("https://www.nuget.org/profiles/weihanli/avatar?imageSize=64");
        var list = Enumerable.Range(1, 5)
            .Select(x => new ImageTest() { Id = x, Image = imageBytes })
            .ToList();
        var excelBytes = list.ToExcelBytes(excelFormat);
        var importResult = ExcelHelper.ToEntityList<ImageTestPicData>(excelBytes, excelFormat);
        Assert.NotNull(importResult);
        Assert.Equal(list.Count, importResult.Count);
        for (var i = 0; i < list.Count; i++)
        {
            Assert.NotNull(importResult[i]);
            var result = importResult[i]!;
            Assert.Equal(list[i].Id, result.Id);
            Assert.NotNull(result.Image);
            Assert.True(list[i].Image.SequenceEqual(result.Image.Data));
            Assert.Equal(PictureType.PNG, result.Image.PictureType);
        }
    }

    [Fact]
    public void DataTableDefaultValueTest()
    {
        var table = new DataTable();
        table.Columns.Add(new DataColumn("Name"));
        table.Columns.Add(new DataColumn("Value"));
        table.Columns.Add(new DataColumn("Description"));
        var row = table.AddNewRow();
        row["Value"] = null;
        row["Description"] = "test";

        Assert.Equal(DBNull.Value, table.Rows[0]["Name"]);
        Assert.Equal(DBNull.Value, table.Rows[0]["Value"]);
        Assert.NotNull(table.Rows[0][0]);
        Assert.Equal("test", table.Rows[0]["Description"]);
    }

    [Theory]
    [ExcelFormatData]
    public void SheetNameTest_ToExcelFile(ExcelFormat excelFormat)
    {
        IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
        {
            Id = i + 1,
            Content = $"content_{i}",
            Title = $"title_{i}",
            PublishedAt = DateTime.UtcNow.AddDays(-i),
            Publisher = $"publisher_{i}"
        }).ToArray();
        var settings = FluentSettings.For<Notice>();
        lock (settings)
        {
            settings.HasSheetSetting(s =>
            {
                s.SheetName = "Test";
            });

            var filePath = $"{Path.GetTempFileName()}.{excelFormat.ToString().ToLower()}";
            list.ToExcelFile(filePath);

            var excel = ExcelHelper.LoadExcel(filePath);
            Assert.Equal("Test", excel.GetSheetAt(0).SheetName);

            settings.HasSheetSetting(s =>
            {
                s.SheetName = "NoticeList";
            });
        }


    }

    [Theory]
    [ExcelFormatData]
    public void SheetNameTest_ToExcelBytes(ExcelFormat excelFormat)
    {
        IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
        {
            Id = i + 1,
            Content = $"content_{i}",
            Title = $"title_{i}",
            PublishedAt = DateTime.UtcNow.AddDays(-i),
            Publisher = $"publisher_{i}"
        }).ToArray();
        var settings = FluentSettings.For<Notice>();
        lock (settings)
        {
            settings.HasSheetSetting(s =>
            {
                s.SheetName = "Test";
            });

            var excelBytes = list.ToExcelBytes(excelFormat);
            var excel = ExcelHelper.LoadExcel(excelBytes, excelFormat);
            Assert.Equal("Test", excel.GetSheetAt(0).SheetName);

            settings.HasSheetSetting(s =>
            {
                s.SheetName = "NoticeList";
            });
        }
    }

    [Theory]
    [ExcelFormatData]
    public void DuplicateColumnTest(ExcelFormat excelFormat)
    {
        var workbook = ExcelHelper.PrepareWorkbook(excelFormat);
        var sheet = workbook.CreateSheet();
        var headerRow = sheet.CreateRow(0);
        headerRow.CreateCell(0).SetCellValue("A");
        headerRow.CreateCell(1).SetCellValue("B");
        headerRow.CreateCell(2).SetCellValue("C");
        headerRow.CreateCell(3).SetCellValue("A");
        headerRow.CreateCell(4).SetCellValue("B");
        headerRow.CreateCell(5).SetCellValue("C");
        var dataRow = sheet.CreateRow(1);
        dataRow.CreateCell(0).SetCellValue("1");
        dataRow.CreateCell(1).SetCellValue("2");
        dataRow.CreateCell(2).SetCellValue("3");
        dataRow.CreateCell(3).SetCellValue("4");
        dataRow.CreateCell(4).SetCellValue("5");
        dataRow.CreateCell(5).SetCellValue("6");
        var dataTable = sheet.ToDataTable();
        Assert.Equal(headerRow.Cells.Count, dataTable.Columns.Count);
        Assert.Equal(1, dataTable.Rows.Count);

        var newWorkbook = ExcelHelper.LoadExcel(dataTable.ToExcelBytes());
        var newSheet = newWorkbook.GetSheetAt(0);
        Assert.Equal(sheet.PhysicalNumberOfRows, newSheet.PhysicalNumberOfRows);
        for (var i = 0; i < sheet.PhysicalNumberOfRows; i++)
        {
            Assert.Equal(sheet.GetRow(i).Cells.Count, newSheet.GetRow(i).Cells.Count);

            for (var j = 0; j < headerRow.Cells.Count; j++)
            {
                Assert.Equal(
                    sheet.GetRow(i).GetCell(j).GetCellValue<string>(),
                    newSheet.GetRow(i).GetCell(j).GetCellValue<string>()
                    );
            }
        }

    }

    [Theory]
    [ExcelFormatData]
    public void ValidatorTest(ExcelFormat excelFormat)
    {
        var list = new List<Job>()
        {
            new()
            {
                Id = 1,
                Name = "test"
            },
            new()
        };
        var bytes = list.ToExcelBytes(excelFormat);
        var result = ExcelHelper.ToEntityListWithValidationResult<Job>(bytes, excelFormat);
        Assert.Equal(list.Count, result.EntityList.Count);
        for (var i = 0; i < list.Count; i++)
        {
            Assert.True(list[i] == result.EntityList[i]);
        }
        Assert.Single(result.ValidationResults);
    }

    [Theory]
    [ExcelFormatData]
    public void ValidatorTest_CustomValidator(ExcelFormat excelFormat)
    {
        var list = new List<Job>()
        {
            new()
            {
                Id = 1,
                Name = "test"
            }
        };
        var validator = new DelegateValidator<Job>(_ => new ValidationResult()
        {
            Valid = false,
            Errors = new Dictionary<string, string[]>() { { "", new[] { "Mock error" } } }
        });
        var bytes = list.ToExcelBytes(excelFormat);
        var result = ExcelHelper.ToEntityListWithValidationResult(bytes, excelFormat, validator: validator);
        Assert.Equal(list.Count, result.EntityList.Count);
        for (var i = 0; i < list.Count; i++)
        {
            Assert.True(list[i] == result.EntityList[i]);
        }
        Assert.Single(result.ValidationResults);
    }

    [Theory]
    [ExcelFormatData]
    public void CellReaderTest(ExcelFormat excelFormat)
    {
        var jobs = new CellReaderTestModel[] { new() { Id = 1, Name = "test" }, new() { Id = 2 }, };
        var bytes = jobs.ToExcelBytes(excelFormat);
        var settings = FluentSettings.For<CellReaderTestModel>();
        settings.Property(x => x.Name)
            .HasCellReader(_ => "CellValue");

        var list = ExcelHelper.ToEntityList<CellReaderTestModel>(bytes, excelFormat);
        Assert.Equal(jobs.Length, list.Count);
        for (var i = 0; i < jobs.Length; i++)
        {
            Assert.NotNull(list[i]);
            var model = list[i];
            Guard.NotNull(model);
            Assert.Equal(jobs[i].Id, model.Id);
            Assert.Equal("CellValue", model.Name);
        }

        settings.Property(x => x.Name)
            .HasCellReader(null);
    }

    [Theory]
    [ExcelFormatData]
    public void CellTypeTest(ExcelFormat excelFormat)
    {
        var workbook = ExcelHelper.PrepareWorkbook(excelFormat);
        var sheet = workbook.CreateSheet();
        var headerRow = sheet.CreateRow(0);
        headerRow.CreateCell(0).SetCellValue("Id");
        headerRow.CreateCell(1).SetCellValue("1234");

        var dataRow = sheet.CreateRow(1);
        dataRow.CreateCell(0).SetCellValue(1);
        dataRow.CreateCell(1).SetCellValue(0.24);

        var cell = dataRow.GetCell(1);
        cell.CellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0%");
        Assert.Equal(CellType.Numeric, cell.CellType);
        Assert.Equal("24%", new DataFormatter().FormatCellValue(cell));

        var excelBytes = workbook.ToExcelBytes();

        var importedWorkbook = ExcelHelper.LoadExcel(excelBytes, excelFormat);
        var importedSheet = importedWorkbook.GetSheetAt(0);
        var row1 = importedSheet.GetRow(1);
        var cell1 = row1.GetCell(1);
        Assert.Equal(cell.CellStyle.DataFormat, cell1.CellStyle.DataFormat);

        Assert.Equal(1, row1.GetCell(0).NumericCellValue.To<int>());
        Assert.Equal("24%", new DataFormatter().FormatCellValue(cell1));
    }

    [Theory]
    [ExcelFormatData]
    public void HeaderCellTypeTest(ExcelFormat excelFormat)
    {
        var workbook = ExcelHelper.PrepareWorkbook(excelFormat);
        var sheet = workbook.CreateSheet();
        var headerRow = sheet.CreateRow(0);
        headerRow.CreateCell(0).SetCellValue("Id");
        var cell = headerRow.CreateCell(1);
        cell.SetCellValue(1234);
        Assert.Equal(CellType.Numeric, cell.CellType);

        var dataRow = sheet.CreateRow(1);
        dataRow.CreateCell(0).SetCellValue(1);
        dataRow.CreateCell(1).SetCellValue("1234");

        var excelBytes = workbook.ToExcelBytes();

        var list = ExcelHelper.ToEntityList<CellFormatTestModel>(excelBytes, excelFormat);
        Assert.Single(list);
        Assert.NotNull(list[0]);
        var entity = Guard.NotNull(list[0]);
        Assert.Equal(1, entity.Id);
        Assert.Equal("1234", entity.Name);
    }

    [Theory]
    [ExcelFormatData]
    public void ChineseDateFormatterTest(ExcelFormat excelFormat)
    {
        FluentSettings.For<ChineseDateFormatter.ChineDateTestModel>()
            .Property(x => x.Date)
            .HasColumnInputFormatter(ChineseDateFormatter.FormatInput)
            .HasColumnOutputFormatter(ChineseDateFormatter.FormatOutput);

        var model = new[]
        {
            new ChineseDateFormatter.ChineDateTestModel() { Date = DateTime.Parse("2022-01-01") }
        };
        var excelBytes = model.ToExcelBytes(excelFormat);
        var list = ExcelHelper.ToEntityList<ChineseDateFormatter.ChineDateTestModel>(excelBytes, excelFormat);
        Assert.Single(list);
        var item = list[0];
        Assert.NotNull(item);
        Guard.NotNull(item);
        Assert.Equal(DateTime.Parse("2022-01-01"), item.Date);
    }

    private sealed class CellFormatTestModel
    {
        public int Id { get; set; }
        [Column("1234")]
        public string? Name { get; set; }
    }

    private sealed record CellReaderTestModel
    {
        public int Id { get; set; }
        public string? Name { get; set; }
    }

    private sealed class ImageTest
    {
        public int Id { get; set; }

        public byte[] Image { get; set; } = null!;
    }

    private sealed class ImageTestPicData
    {
        public int Id { get; set; }

        public IPictureData Image { get; set; } = null!;
    }

    private sealed class ChineseDateFormatter
    {
        public sealed class ChineDateTestModel
        {
            public DateTime Date { get; set; }
        }

        public static DateTime FormatInput(string? input)
        {
            if (DateTimeUtils.TransStrToDateTime(input, out var dt))
            {
                return dt;
            }
            throw new ArgumentException("Invalid date input");
        }

        public static string FormatOutput(DateTime input)
        {
            return "二〇二二年一月一日";
        }
    }
}

// http://luoma.pro/Content/Detail/671?parentId=1
public static class DateTimeUtils
{
    /// <summary>
    /// 字符串日期转 DateTime  
    /// </summary>
    /// <param name="str">字符串日期</param>
    /// <param name="dt">转换成功赋值</param>
    /// <returns>转换成功返回 true</returns>
    public static bool TransStrToDateTime(string? str, out DateTime dt)
    {
        dt = default;
        if (str.IsNullOrEmpty())
            return false;

        //第一次转换
        if (DateTime.TryParse(str, out dt))
        {
            return true;
        }
        //第二次转换
        string[] format = new string[]
        {
            "yyyyMMdd",
            "yyyyMdHHmmss",
            "yyyyMMddHHmmss",
            "yyyy-M-d",
            "yyyy-MM-dd",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy/M/d",
            "yyyy/MM/dd",
            "yyyy/MM/dd HH:mm:ss",
            "yyyy.M.d",
            "yyyy.MM.dd",
            "yyyy.MM.dd HH:mm:ss",
            "yyyy年M月d日",
            "yyyy年MM月dd日",
            "yyyy年MM月dd日HH:mm:ss",
            "yyyy年MM月dd日 HH时mm分ss秒"
        };
        if (DateTime.TryParseExact(str, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
        {
            return true;
        }
        //第三次转换
        try
        {
            if (Regex.IsMatch(str, "^(零|〇|一|二|三|四|五|六|七|八|九|十){2,4}年((正|一|二|三|四|五|六|七|八|九|十|十一|十二)月((一|二|三|四|五|六|七|八|九|十){1,3}(日)?)?)?$"))
            {
                var match = Regex.Match(str, @"^(.+)年(.+)月(.+)日$");
                if (match.Success)
                {
                    int year = GetYear(match.Groups[1].Value);
                    int month = GetMonth(match.Groups[2].Value);
                    long dayL = ParseCnToInt(match.Groups[3].Value);
                    dt = new DateTime(year, month, int.Parse(dayL.ToString()));
                    return true;
                }
            }
        }
        catch
        {
            return false;
        }
        return false;
    }
    /// <summary>
    /// 使用正则表达式判断是否为日期
    /// </summary>
    /// <param name="str">日期格式字符串</param>
    /// <returns>是日期格式字符串返回 true</returns>
    public static bool IsDateTime(string str)
    {
        bool isDateTime;
        // yyyy/MM/dd  - 年月日数字
        if (Regex.IsMatch(str, "^(?<year>\\d{2,4})/(?<month>\\d{1,2})/(?<day>\\d{1,2})$"))
            isDateTime = true;
        // yyyy-MM-dd - 年月日数字  
        else if (Regex.IsMatch(str, "^(?<year>\\d{2,4})-(?<month>\\d{1,2})-(?<day>\\d{1,2})$"))
            isDateTime = true;
        // yyyy.MM.dd - 年月日数字  
        else if (Regex.IsMatch(str, "^(?<year>\\d{2,4})[.](?<month>\\d{1,2})[.](?<day>\\d{1,2})$"))
            isDateTime = true;
        // yyyy年MM月dd日 - 年月日数字  
        else if (Regex.IsMatch(str, "^((?<year>\\d{2,4})年)?(?<month>\\d{1,2})月((?<day>\\d{1,2})日)?$"))
            isDateTime = true;
        // yyyy年MM月dd日  - 年月日中文 
        else if (Regex.IsMatch(str, "^(零|〇|一|二|三|四|五|六|七|八|九|十){2,4}年((正|一|二|三|四|五|六|七|八|九|十|十一|十二)月((一|二|三|四|五|六|七|八|九|十){1,3}(日)?)?)?$"))
            isDateTime = true;
        // yyyy年MM月dd日  - 年(数字)，月(中文)，日(中文)
        //else if (Regex.IsMatch(str, "^((?<year>\\d{2,4})年)?(正|一|二|三|四|五|六|七|八|九|十|十一|十二)月((一|二|三|四|五|六|七|八|九|十){1,3}日)?$"))
        //    isDateTime = true;
        // yyyy年  
        //else if (Regex.IsMatch(str, "^(?<year>\\d{2,4})年$"))  
        //    isDateTime = true;  
        // 农历1  
        //else if (Regex.IsMatch(str, "^(甲|乙|丙|丁|戊|己|庚|辛|壬|癸)(子|丑|寅|卯|辰|巳|午|未|申|酉|戌|亥)年((正|一|二|三|四|五|六|七|八|九|十|十一|十二)月((一|二|三|四|五|六|七|八|九|十){1,3}(日)?)?)?$"))
        //    isDateTime = true;
        //// 农历2  
        //else if (Regex.IsMatch(str, "^((甲|乙|丙|丁|戊|己|庚|辛|壬|癸)(子|丑|寅|卯|辰|巳|午|未|申|酉|戌|亥)年)?(正|一|二|三|四|五|六|七|八|九|十|十一|十二)月初(一|二|三|四|五|六|七|八|九|十)$"))
        //    isDateTime = true;
        //// XX时XX分XX秒  
        //else if (Regex.IsMatch(str, "^(?<hour>\\d{1,2})(时|点)(?<minute>\\d{1,2})分((?<second>\\d{1,2})秒)?$"))
        //    isDateTime = true;
        //// XX时XX分XX秒  
        //else if (Regex.IsMatch(str, "^((零|一|二|三|四|五|六|七|八|九|十){1,3})(时|点)((零|一|二|三|四|五|六|七|八|九|十){1,3})分(((零|一|二|三|四|五|六|七|八|九|十){1,3})秒)?$"))
        //    isDateTime = true;
        //// XX分XX秒  
        //else if (Regex.IsMatch(str, "^(?<minute>\\d{1,2})分(?<second>\\d{1,2})秒$"))
        //    isDateTime = true;
        //// XX分XX秒  
        //else if (Regex.IsMatch(str, "^((零|一|二|三|四|五|六|七|八|九|十){1,3})分((零|一|二|三|四|五|六|七|八|九|十){1,3})秒$"))
        //    isDateTime = true;
        //// XX时  
        //else if (Regex.IsMatch(str, "\\b(?<hour>\\d{1,2})(时|点钟)\\b"))
        //    isDateTime = true;
        else
            isDateTime = false;
        return isDateTime;
    }
    #region 年月获取
    /// <summary>
    /// 获取年份
    /// </summary>
    /// <param name="str">年份</param>
    /// <returns>数字年份</returns>
    public static int GetYear(string str)
    {
        var strNumber = "";
        foreach (var item in str)
        {
            switch (item.ToString())
            {
                case "零":
                case "〇":
                    strNumber += "0";
                    break;
                case "一":
                    strNumber += "1";
                    break;
                case "二":
                    strNumber += "2";
                    break;
                case "三":
                    strNumber += "3";
                    break;
                case "四":
                    strNumber += "4";
                    break;
                case "五":
                    strNumber += "5";
                    break;
                case "六":
                    strNumber += "6";
                    break;
                case "七":
                    strNumber += "7";
                    break;
                case "八":
                    strNumber += "8";
                    break;
                case "九":
                    strNumber += "9";
                    break;
                case "十":
                    strNumber += "10";
                    break;
            }
        }
        int.TryParse(strNumber, out var number);
        return number;
    }
    /// <summary>
    ///获取月份
    /// </summary>
    /// <param name="str">月份</param>
    /// <returns>数字月份</returns>
    public static int GetMonth(string str)
    {
        var strNumber = "";
        switch (str)
        {
            case "一":
            case "正":
                strNumber += "1";
                break;
            case "二":
                strNumber += "2";
                break;
            case "三":
                strNumber += "3";
                break;
            case "四":
                strNumber += "4";
                break;
            case "五":
                strNumber += "5";
                break;
            case "六":
                strNumber += "6";
                break;
            case "七":
                strNumber += "7";
                break;
            case "八":
                strNumber += "8";
                break;
            case "九":
                strNumber += "9";
                break;
            case "十":
                strNumber += "10";
                break;
            case "十一":
                strNumber += "11";
                break;
            case "十二":
                strNumber += "12";
                break;
        }
        int.TryParse(strNumber, out var number);
        return number;
    }
    #endregion
    #region 中文数字和阿拉伯数字转换
    /// <summary>
    /// 阿拉伯数字转换成中文数字
    /// </summary>
    /// <param name="x"></param>
    /// <returns></returns>
    public static string NumToChinese(string x)
    {
        string[] pArrayNum = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
        //为数字位数建立一个位数组
        string[] pArrayDigit = { "", "十", "百", "千" };
        //为数字单位建立一个单位数组
        string[] pArrayUnits = { "", "万", "亿", "万亿" };
        var pStrReturnValue = ""; //返回值
        var finger = 0; //字符位置指针
        var pIntM = x.Length % 4; //取模
        int pIntK;
        if (pIntM > 0)
            pIntK = x.Length / 4 + 1;
        else
            pIntK = x.Length / 4;
        //外层循环,四位一组,每组最后加上单位: ",万亿,",",亿,",",万,"
        for (var i = pIntK; i > 0; i--)
        {
            var pIntL = 4;
            if (i == pIntK && pIntM != 0)
                pIntL = pIntM;
            //得到一组四位数
            var four = x.Substring(finger, pIntL);
            var pIntL1 = four.Length;
            //内层循环在该组中的每一位数上循环
            for (var j = 0; j < pIntL1; j++)
            {
                //处理组中的每一位数加上所在的位
                var n = Convert.ToInt32(four.Substring(j, 1));
                if (n == 0)
                {
                    if (j < pIntL1 - 1 && Convert.ToInt32(four.Substring(j + 1, 1)) > 0 && !pStrReturnValue.EndsWith(pArrayNum[n]))
                        pStrReturnValue += pArrayNum[n];
                }
                else
                {
                    if (!(n == 1 && (pStrReturnValue.EndsWith(pArrayNum[0]) | pStrReturnValue.Length == 0) && j == pIntL1 - 2))
                        pStrReturnValue += pArrayNum[n];
                    pStrReturnValue += pArrayDigit[pIntL1 - j - 1];
                }
            }
            finger += pIntL;
            //每组最后加上一个单位:",万,",",亿," 等
            if (i < pIntK) //如果不是最高位的一组
            {
                if (Convert.ToInt32(four) != 0)
                    //如果所有4位不全是0则加上单位",万,",",亿,"等
                    pStrReturnValue += pArrayUnits[i - 1];
            }
            else
            {
                //处理最高位的一组,最后必须加上单位
                pStrReturnValue += pArrayUnits[i - 1];
            }
        }
        return pStrReturnValue;
    }
    /// <summary>
    /// 转换数字
    /// </summary>
    public static long CharToNumber(char c)
    {
        switch (c)
        {
            case '一': return 1;
            case '二': return 2;
            case '三': return 3;
            case '四': return 4;
            case '五': return 5;
            case '六': return 6;
            case '七': return 7;
            case '八': return 8;
            case '九': return 9;
            case '零': return 0;
            default: return -1;
        }
    }
    /// <summary>
    /// 转换单位
    /// </summary>
    public static long CharToUnit(char c)
    {
        switch (c)
        {
            case '十': return 10;
            case '百': return 100;
            case '千': return 1000;
            case '万': return 10000;
            case '亿': return 100000000;
            default: return 1;
        }
    }
    /// <summary>
    /// 将中文数字转换阿拉伯数字
    /// </summary>
    /// <param name="cnum">汉字数字</param>
    /// <returns>长整型阿拉伯数字</returns>
    public static long ParseCnToInt(string cnum)
    {
        cnum = Regex.Replace(cnum, "\\s+", "");
        long firstUnit = 1;//一级单位
        long secondUnit = 1;//二级单位
        long result = 0;//结果
        for (var i = cnum.Length - 1; i > -1; --i)//从低到高位依次处理
        {
            var tmpUnit = CharToUnit(cnum[i]);//临时单位变量
            if (tmpUnit > firstUnit)//判断此位是数字还是单位
            {
                firstUnit = tmpUnit;//是的话就赋值,以备下次循环使用
                secondUnit = 1;
                if (i == 0)//处理如果是"十","十一"这样的开头的
                {
                    result += firstUnit * secondUnit;
                }
                continue;//结束本次循环
            }
            if (tmpUnit > secondUnit)
            {
                secondUnit = tmpUnit;
                continue;
            }
            result += firstUnit * secondUnit * CharToNumber(cnum[i]);//如果是数字,则和单位想乘然后存到结果里
        }
        return result;
    }
    #endregion
}
