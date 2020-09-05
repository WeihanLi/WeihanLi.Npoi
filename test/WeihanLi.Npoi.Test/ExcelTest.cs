using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using WeihanLi.Document;
using WeihanLi.Document.Excel;
using WeihanLi.Extensions;
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
            var noticeSetting = FluentSettings.ExcelSettingsFor<Notice>();
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
            var noticeSetting = FluentSettings.ExcelSettingsFor<Notice>();
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

            var noticeSetting = FluentSettings.ExcelSettingsFor<Notice>();
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
            var noticeSetting = FluentSettings.ExcelSettingsFor<Notice>();
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

            var noticeSetting = FluentSettings.ExcelSettingsFor<Notice>();
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

            var settings = FluentSettings.ExcelSettingsFor<Notice>();
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

            var settings = FluentSettings.ExcelSettingsFor<Notice>();
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

            var settings = FluentSettings.ExcelSettingsFor<Notice>();
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
    }
}
