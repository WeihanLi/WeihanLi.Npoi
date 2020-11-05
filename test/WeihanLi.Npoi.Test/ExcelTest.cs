using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Attributes;
using WeihanLi.Npoi.Test.Models;
using Xunit;

namespace WeihanLi.Npoi.Test
{
    public class ExcelTest : TestBase
    {
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
                        Assert.Equal(list[i].Id, importedList[i].Id);
                        Assert.Equal(list[i].Title, importedList[i].Title);
                        Assert.Equal(list[i].Content, importedList[i].Content);
                        Assert.Equal(list[i].Publisher, importedList[i].Publisher);
                        Assert.Equal(list[i].PublishedAt.ToStandardTimeString(), importedList[i].PublishedAt.ToStandardTimeString());
                    }
                }
            }
        }

        [Theory]
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Xlsx)]
        public void BasicImportExportTestWithEmptyValue(ExcelFormat excelFormat)
        {
            var list = new List<Notice>();
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
                        Assert.Equal(list[i].Id, importedList[i].Id);
                        Assert.Equal(list[i].Title, importedList[i].Title);
                        Assert.Equal(list[i].Content, importedList[i].Content);
                        Assert.Equal(list[i].Publisher, importedList[i].Publisher);
                        Assert.Equal(list[i].PublishedAt.ToStandardTimeString(), importedList[i].PublishedAt.ToStandardTimeString());
                    }
                }
            }
        }

        [Theory]
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Xlsx)]
        public void BasicImportExportWithoutHeaderTest(ExcelFormat excelFormat)
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
                        Assert.Equal(list[i].Id, importedList[i].Id);
                        Assert.Equal(list[i].Title, importedList[i].Title);
                        Assert.Equal(list[i].Content, importedList[i].Content);
                        Assert.Equal(list[i].Publisher, importedList[i].Publisher);
                        Assert.Equal(list[i].PublishedAt.ToStandardTimeString(), importedList[i].PublishedAt.ToStandardTimeString());
                    }
                }

                noticeSetting.HasSheetConfiguration(0, "test", 1);
            }
        }

        [Theory]
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Xlsx)]
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
                        Assert.Equal(list[i].Id, importedList[i].Id);
                        Assert.Equal(list[i].Title, importedList[i].Title);
                        Assert.Equal(list[i].Content, importedList[i].Content);
                        Assert.Equal(list[i].Publisher, importedList[i].Publisher);
                        Assert.Equal(list[i].PublishedAt.ToStandardTimeString(), importedList[i].PublishedAt.ToStandardTimeString());
                    }
                }

                noticeSetting.Property(_ => _.Publisher)
                    .HasColumnIndex(3);
                noticeSetting.Property(_ => _.PublishedAt)
                    .HasColumnIndex(4);
            }
        }

        [Theory]
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Xlsx)]
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
                    .HasOutputFormatter((x, val) => $"{x.Id}...")
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
                        Assert.Equal(list[i].Id, importedList[i].Id);
                        Assert.Equal(list[i].Title, importedList[i].Title);
                        Assert.Equal(list[i].Content, importedList[i].Content);
                        Assert.Equal(list[i].Publisher, importedList[i].Publisher);
                        Assert.Equal(list[i].PublishedAt.ToStandardTimeString(), importedList[i].PublishedAt.ToStandardTimeString());
                    }
                }
            }
        }

        [Theory]
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Xlsx)]
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
                        // Assert.Equal(list[i].Id, importedList[i].Id);
                        Assert.Equal(list[i].Title, importedList[i].Title);
                        Assert.Equal(list[i].Content, importedList[i].Content);
                        Assert.Equal(list[i].Publisher, importedList[i].Publisher);
                        Assert.Equal(list[i].PublishedAt.ToStandardTimeString(), importedList[i].PublishedAt.ToStandardTimeString());
                    }
                }

                settings.Property(_ => _.Id)
                    .Ignored(false)
                    .HasColumnIndex(0);
            }
        }

        [Theory]
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Xlsx)]
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
                        Assert.Equal(list[i].Id, importedList[i].Id);
                        Assert.Equal(list[i].Title + "_Test", importedList[i].Title);
                        Assert.Equal(list[i].Content, importedList[i].Content);
                        Assert.Equal(list[i].Publisher, importedList[i].Publisher);
                        Assert.Equal(list[i].PublishedAt.ToStandardTimeString(), importedList[i].PublishedAt.ToStandardTimeString());
                    }
                }

                settings.Property(_ => _.Title).HasColumnInputFormatter(null);
            }
        }

        [Theory]
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Xlsx)]
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
                    .HasColumnInputFormatter(x => Convert.ToInt32(x.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries)[0]));
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
                        Assert.Equal(list[i].PublishedAt.ToStandardTimeString(), importedList[i].PublishedAt.ToStandardTimeString());
                    }
                }

                settings.Property(x => x.Id)
                    .HasColumnOutputFormatter(null)
                    .HasColumnInputFormatter(null);
            }
        }

        [Theory]
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Xlsx)]
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
                settings.WithDataValidation(x => x.Id > 5);

                var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes, excelFormat);
                Assert.Equal(list.Count(x => x.Id > 5), importedList.Count);

                int i = 0, k = 0;
                while (list[k].Id != importedList[i].Id)
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
                        Assert.Equal(list[k].Id, importedList[i].Id);
                        Assert.Equal(list[k].Title, importedList[i].Title);
                        Assert.Equal(list[k].Content, importedList[i].Content);
                        Assert.Equal(list[k].Publisher, importedList[i].Publisher);
                        Assert.Equal(list[k].PublishedAt.ToStandardTimeString(), importedList[i].PublishedAt.ToStandardTimeString());
                    }
                }

                settings.WithDataValidation(null);
            }
        }

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

        [Theory]
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Xlsx)]
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
                row.ItemArray = new object[]
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
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Xlsx)]
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
            Assert.Equal(1, list[0].Num1);
            Assert.Equal(2, list[0].Num2);
            Assert.Equal(3, list[0].Sum);
        }

        // ReSharper disable UnusedAutoPropertyAccessor.Local
        private class ExcelFormulaTestModel
        {
            public int Num1 { get; set; }
            public int Num2 { get; set; }

            public int Sum { get; set; }
        }

        [Theory]
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Xlsx)]
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
                        Assert.Equal(list[i].Id, importedList[i].Id);
                        Assert.Null(importedList[i].Title);
                        Assert.Null(importedList[i].Content);
                        Assert.Null(importedList[i].Publisher);
                        Assert.Equal(default(DateTime).ToStandardTimeString(), importedList[i].PublishedAt.ToStandardTimeString());
                    }
                }

                settings.HasSheetSetting(setting =>
                {
                    setting.CellFilter = null;
                });
            }
        }

        [Theory]
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Xlsx)]
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
                Assert.Equal(list[i].Id, importedList[i].Id);
                Assert.Equal(list[i].Name, importedList[i].Name);
                Assert.Null(importedList[i].Description);
            }
        }

        [Sheet(SheetName = "test", AutoColumnWidthEnabled = true, StartColumnIndex = 0, EndColumnIndex = 1)]
        private class CellFilterAttributeTest
        {
            [Column(Index = 0)]
            public int Id { get; set; }

            [Column(Index = 1)]
            public string Name { get; set; }

            [Column(Index = 2)]
            public string Description { get; set; }
        }
    }
}
