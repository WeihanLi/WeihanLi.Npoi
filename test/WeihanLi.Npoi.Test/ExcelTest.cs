using System;
using System.Collections.Generic;
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
}
