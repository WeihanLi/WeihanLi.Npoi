using System;
using System.Collections.Generic;
using System.Linq;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Test.Models;
using Xunit;

namespace WeihanLi.Npoi.Test
{
    public class ExcelTest : TestBase
    {
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
            list.Add(new Notice() { Title = "nnnn" });
            list.Add(null);
            var noticeSetting = ExcelHelper.SettingFor<Notice>();
            lock (noticeSetting)
            {
                var excelBytes = list.ToExcelBytes();

                var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes);
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

        [Fact]
        public void BasicImportExportWithoutHeaderTest()
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

            var noticeSetting = ExcelHelper.SettingFor<Notice>();
            lock (noticeSetting)
            {
                noticeSetting.HasSheetConfiguration(0, "test", 0);

                var excelBytes = list.ToExcelBytes();

                var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes);
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

        [Fact]
        public void ImportWithNotSpecificColumnIndex()
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
            var noticeSetting = ExcelHelper.SettingFor<Notice>();
            lock (noticeSetting)
            {
                var excelBytes = list.ToExcelBytes();

                noticeSetting.Property(_ => _.Publisher)
                    .HasColumnIndex(4);
                noticeSetting.Property(_ => _.PublishedAt)
                    .HasColumnIndex(3);

                var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes);
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

        [Fact]
        public void HiddenPropertyTest()
        {
            IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
            {
                Id = i + 1,
                Content = $"content_{i}",
                Title = $"title_{i}",
                PublishedAt = DateTime.UtcNow.AddDays(-i),
                Publisher = $"publisher_{i}"
            }).ToArray();

            var noticeSetting = ExcelHelper.SettingFor<Notice>();
            lock (noticeSetting)
            {
                noticeSetting.Property<string>("HiddenProperty")
                    .HasOutputFormatter((x, val) => $"{x.Id}...")
                    ;

                var excelBytes = list.ToExcelBytes();
                // list.ToExcelFile($"{Directory.GetCurrentDirectory()}/output.xlsx");

                var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes);
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

        [Fact]
        public void IgnoreInheritPropertyTest()
        {
            IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
            {
                Id = i + 1,
                Content = $"content_{i}",
                Title = $"title_{i}",
                PublishedAt = DateTime.UtcNow.AddDays(-i),
                Publisher = $"publisher_{i}"
            }).ToArray();

            var settings = ExcelHelper.SettingFor<Notice>();
            lock (settings)
            {
                settings.Property(x => x.Id).Ignored();

                var excelBytes = list.ToExcelBytes();
                // list.ToExcelFile($"{Directory.GetCurrentDirectory()}/ttt.xls");
                var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes);
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

        [Fact]
        public void ColumnInputFormatterTest()
        {
            IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
            {
                Id = i + 1,
                Content = $"content_{i}",
                Title = $"title_{i}",
                PublishedAt = DateTime.UtcNow.AddDays(-i),
                Publisher = $"publisher_{i}"
            }).ToArray();

            var excelBytes = list.ToExcelBytes();

            var settings = ExcelHelper.SettingFor<Notice>();
            lock (settings)
            {
                settings.Property(x => x.Title).HasColumnInputFormatter(x => $"{x}_Test");

                var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes);
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

        [Fact]
        public void InputOutputColumnFormatterTest()
        {
            IReadOnlyList<Notice> list = Enumerable.Range(0, 10).Select(i => new Notice()
            {
                Id = i + 1,
                Content = $"content_{i}",
                Title = $"title_{i}",
                PublishedAt = DateTime.UtcNow.AddDays(-i),
                Publisher = $"publisher_{i}"
            }).ToArray();

            var settings = ExcelHelper.SettingFor<Notice>();
            lock (settings)
            {
                settings.Property(x => x.Id)
                    .HasColumnOutputFormatter(x => $"{x}_Test")
                    .HasColumnInputFormatter(x => Convert.ToInt32(x.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries)[0]));
                var excelBytes = list.ToExcelBytes();

                var importedList = ExcelHelper.ToEntityList<Notice>(excelBytes);
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
    }
}
