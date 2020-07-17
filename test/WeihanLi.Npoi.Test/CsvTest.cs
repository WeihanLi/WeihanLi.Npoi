using System;
using System.Collections.Generic;
using System.Data;
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
            var noticeSetting = FluentSettings.For<Notice>();
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
            var noticeSetting = FluentSettings.For<Notice>();
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

        [Theory]
        [InlineData("\"XXXXX\"")]
        [InlineData("XXX")]
        [InlineData("\"X,XXX\"")]
        [InlineData("XX\"X")]
        [InlineData("XX\"\"X")]
        [InlineData("\"dd\"\"d,1\"")]
        [InlineData("ddd\nccc")]
        [InlineData("ddd\r\nccc")]
        [InlineData(@"bbb
        ccc")]
        [InlineData("")]
        public void ParseCsvLineTest(string str)
        {
            var data = new object[] { 1, "tom", 33, str };
            var lineData = string.Join(CsvHelper.CsvSeparatorCharacter, data);
            var cols = CsvHelper.ParseLine(lineData);
            Assert.Equal(data.Length, cols.Count);

            for (var i = 0; i < cols.Count; i++)
            {
                Assert.Equal(TrimQuotes(data[i]?.ToString()), cols[i]);
            }
        }

        [Fact]
        public void GetCsvTextTest()
        {
            var text = Enumerable.Range(1, 5)
                .GetCsvText(false);

            var expected = Enumerable.Range(1, 5)
                .StringJoin(Environment.NewLine) + Environment.NewLine;
            Assert.Equal(expected, text);
        }

        private static string TrimQuotes(string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return string.Empty;
            }
            //

            if (str[0] == CsvHelper.CsvQuoteCharacter)
            {
                return str.Substring(1, str.Length - 2).Replace("\"\"", "\"");
            }

            return str;
        }
    }
}
