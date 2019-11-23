using System;
using System.Collections.Generic;
using System.Linq;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Test.Models;
using Xunit;

namespace WeihanLi.Npoi.Test
{
    public class CsvTest : TestBase
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
            list.Add(new Notice()
            {
                Id = 11,
                Content = $"content",
                Title = $"title",
                PublishedAt = DateTime.UtcNow.AddDays(1),
            });
            var noticeSetting = ExcelHelper.SettingFor<Notice>();
            lock (noticeSetting)
            {
                var csvBytes = list.ToCsvBytes();
                var importedList = CsvHelper.ToEntityList<Notice>(csvBytes);
                Assert.Equal(list.Count, importedList.Count);
                for (var i = 0; i < list.Count; i++)
                {
                    Assert.Equal(list[i].Id, importedList[i].Id);
                    Assert.Equal(list[i].Title ?? "", importedList[i].Title);
                    Assert.Equal(list[i].Content ?? "", importedList[i].Content);
                    Assert.Equal(list[i].Publisher ?? "", importedList[i].Publisher);
                    Assert.Equal(list[i].PublishedAt.ToStandardTimeString(), importedList[i].PublishedAt.ToStandardTimeString());
                }
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
                var excelBytes = list.ToCsvBytes();

                noticeSetting.Property(_ => _.Publisher)
                    .HasColumnIndex(4);
                noticeSetting.Property(_ => _.PublishedAt)
                    .HasColumnIndex(3);

                var importedList = CsvHelper.ToEntityList<Notice>(excelBytes);
                Assert.Equal(list.Count, importedList.Count);
                for (var i = 0; i < list.Count; i++)
                {
                    Assert.Equal(list[i].Id, importedList[i].Id);
                    Assert.Equal(list[i].Title ?? "", importedList[i].Title);
                    Assert.Equal(list[i].Content ?? "", importedList[i].Content);
                    Assert.Equal(list[i].Publisher ?? "", importedList[i].Publisher);
                    Assert.Equal(list[i].PublishedAt.ToStandardTimeString(), importedList[i].PublishedAt.ToStandardTimeString());
                }

                noticeSetting.Property(_ => _.Publisher)
                    .HasColumnIndex(3);
                noticeSetting.Property(_ => _.PublishedAt)
                    .HasColumnIndex(4);
            }
        }
    }
}
