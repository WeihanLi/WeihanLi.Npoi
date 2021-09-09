using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Attributes;
using WeihanLi.Npoi.Test.Models;
using Xunit;

namespace WeihanLi.Npoi.Test
{
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
                    if (list[i] == null)
                    {
                        Assert.Null(importedList[i]);
                    }
                    else
                    {
                        Assert.NotNull(importedList[i]);
                        var sourceItem = list[i]!;
                        var item = importedList[i]!;
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
                    if (list[i] == null)
                    {
                        Assert.Null(importedList[i]);
                    }
                    else
                    {
                        Assert.NotNull(importedList[i]);
                        var item = importedList[i]!;
                        Assert.Equal(list[i].Id, item.Id);
                        Assert.Equal(list[i].Title + "_Test", item.Title);
                        Assert.Equal(list[i].Content, item.Content);
                        Assert.Equal(list[i].Publisher, item.Publisher);
                        Assert.Equal(list[i].PublishedAt.ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());
                    }
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
                settings.WithDataValidation(x => x?.Id > 5);

                var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes, excelFormat);
                Assert.Equal(list.Count(x => x.Id > 5), importedList.Count);

                int i = 0, k = 0;
                while (list[k].Id != importedList[i]?.Id)
                {
                    k++;
                }

                for (; i < importedList.Count; i++, k++)
                {
                    if (list[k] == null)
                    {
                        Assert.Null(importedList[i]);
                    }
                    else
                    {
                        Assert.NotNull(importedList[i]);
                        var sourceItem = list[k]!;
                        var item = importedList[i]!;
                        Assert.Equal(sourceItem.Id, item.Id);
                        Assert.Equal(sourceItem.Title, item.Title);
                        Assert.Equal(sourceItem.Content, item.Content);
                        Assert.Equal(sourceItem.Publisher, item.Publisher);
                        Assert.Equal(sourceItem.PublishedAt.ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());
                    }
                }

                settings.WithDataValidation(null);
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
                    if (list[i] == null)
                    {
                        Assert.Null(importedList[i]);
                    }
                    else
                    {
                        Assert.NotNull(importedList[i]);
                        var item = importedList[i]!;

                        Assert.Equal(list[i].Id, item.Id);
                        Assert.Null(item.Title);
                        Assert.Null(item.Content);
                        Assert.Null(item.Publisher);
                        Assert.Equal(default(DateTime).ToStandardTimeString(), item.PublishedAt.ToStandardTimeString());
                    }
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
        [InlineData(@"TestData\EmptyColumns\emptyColumns.xls", ExcelFormat.Xls)]
        [InlineData(@"TestData\EmptyColumns\emptyColumns.xlsx", ExcelFormat.Xlsx)]
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

            AssertDataTable(importedData, dt);
        }

        [Theory]
        [InlineData(@"TestData\NonStringColumns\nonStringColumns.xls", ExcelFormat.Xls)]
        [InlineData(@"TestData\NonStringColumns\nonStringColumns.xlsx", ExcelFormat.Xlsx)]
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

            AssertDataTable(importedData, dt);
        }

        [Theory]
        [InlineData(@"TestData\EmptyRows\emptyRows.xls", ExcelFormat.Xls)]
        [InlineData(@"TestData\EmptyRows\emptyRows.xlsx", ExcelFormat.Xlsx)]
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

            AssertDataTable(importedData, dt);
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

        private static void AssertDataTable(DataTable actual, DataTable expected)
        {
            // Check columns
            for (var headerIndex = 0; headerIndex < expected.Columns.Count; headerIndex++)
            {
                var expectedValue = expected.Columns[headerIndex]?.ToString();
                var excelValue = actual.Columns[headerIndex].ToString();

                // "TRUE" from header column is translated to "True".
                // I don't know how to load display value of boolean, therefore I ignore letter casing.
                Assert.Equal(expectedValue, excelValue, ignoreCase: true);
            }

            // Check rows
            for (var rowIndex = 0; rowIndex < expected.Rows.Count; rowIndex++)
            {
                for (var colIndex = 0; colIndex < expected.Rows[rowIndex].ItemArray.Length; colIndex++)
                {
                    var expectedValue = expected.Rows[rowIndex].ItemArray[colIndex]?.ToString();
                    var excelValue = actual.Rows[rowIndex][colIndex].ToString();
                    Assert.Equal(expectedValue, excelValue);
                }
            }
        }

        private class ImageTest
        {
            public int Id { get; set; }

            public byte[] Image { get; set; } = null!;
        }

        private class ImageTestPicData
        {
            public int Id { get; set; }

            public IPictureData Image { get; set; } = null!;
        }
    }
}
