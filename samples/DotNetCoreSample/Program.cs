using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;
using WeihanLi.Npoi;
using WeihanLi.Npoi.Attributes;

// ReSharper disable All
namespace DotNetCoreSample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            FluentSettingsForExcel();
            var tempDirPath = $@"{Environment.GetEnvironmentVariable("USERPROFILE")}\Desktop\temp\test";

            //FluentSettings.For<ppDto>()
            //    .HasSheetSetting(sheet =>
            //    {
            //        sheet.CellFilter = cell => cell.ColumnIndex <= 10;
            //    });
            //var tempExcelPath = Path.Combine(tempDirPath, "testdata.xlsx");
            //var t_list = ExcelHelper.ToEntityList<ppDto>(tempExcelPath);
            //var tempTable = ExcelHelper.ToDataTable(tempExcelPath);

            //using (var conn = new SqlConnection("server=.;uid=liweihan;pwd=Admin888;database=Reservation"))
            //{
            //    var list = conn.Select<TestEntity>(@"SELECT * FROM [Reservation].[dbo].[tabSystemSettings]").ToArray();
            //    list.ToExcelFile(ApplicationHelper.MapPath("test.xlsx"));
            //}

            //var entityList = ExcelHelper.ToEntityList<TestEntity>(ApplicationHelper.MapPath("test.xlsx"));

            //Console.WriteLine("Success!");

            //var mapping = ExcelHelper.ToEntityList<ProductPriceMapping>($@"{Environment.GetEnvironmentVariable("USERPROFILE")}\Desktop\temp\tempFiles\mapping.xlsx");

            //var mappingTemp = ExcelHelper.ToEntityList<ProductPriceMapping>($@"{Environment.GetEnvironmentVariable("USERPROFILE")}\Desktop\temp\tempFiles\mapping_temp.xlsx");

            //Console.WriteLine($"-----normal({mapping.Count}【{mapping.Select(_ => _.Pid).Distinct().Count()}】)----");
            //foreach (var shop in mapping.GroupBy(_ => _.ShopCode).OrderBy(_ => _.Key))
            //{
            //    Console.WriteLine($"{shop.Key}---{shop.Count()}---distinct pid count:{shop.Select(_ => _.Pid).Distinct().Count()}");
            //}

            //Console.WriteLine($"-----temp({mappingTemp.Count}【{mappingTemp.Select(_ => _.Pid).Distinct().Count()}】)----");
            //foreach (var shop in mappingTemp.GroupBy(_ => _.ShopCode).OrderBy(_ => _.Key))
            //{
            //    Console.WriteLine($"{shop.Key}---{shop.Count()}---distinct pid count:{shop.Select(_ => _.Pid).Distinct().Count()}");
            //}

            //Console.WriteLine("Press Enter to continue...");
            //Console.ReadLine();
            var list2 = new List<TestEntity2?>();
            list2.Add(null);
            for (var i = 0; i < 100_000; i++)
            {
                list2.Add(new TestEntity2
                {
                    Id = i + 1,
                    Title = $"Title_{i}",
                    Description = $"{Enumerable.Range(1, 200).StringJoin(",")}__{i}",
                });
            }
            list2.Add(new TestEntity2()
            {
                Id = 999,
                Title = $"{Enumerable.Repeat(1, 10).StringJoin(",")}",
                Description = null
            });
            var watch = Stopwatch.StartNew();
            list2.ToExcelFile($@"{tempDirPath}\testEntity2.xls");
            watch.Stop();
            Console.WriteLine($"ElapsedMilliseconds: {watch.ElapsedMilliseconds}ms");
            //var listTemp = ExcelHelper.ToEntityList<TestEntity2>($@"{tempDirPath}\testEntity2.xlsx");
            //var dataTableTemp = ExcelHelper.ToDataTable($@"{tempDirPath}\testEntity2.xlsx");

            Console.WriteLine("Press Enter to continue...");
            Console.ReadLine();

            var entities = new List<TestEntity>()
            {
                new TestEntity()
                {
                    PKID = 1,
                    SettingId = Guid.NewGuid(),
                    SettingName = "Setting1",
                    SettingValue = "Value1",
                    DisplayName = "dd\"d,1"
                },
                new TestEntity()
                {
                    PKID=2,
                    SettingId = Guid.NewGuid(),
                    SettingName = "Setting2",
                    SettingValue = "Value2",
                    Enabled = true,
                    CreatedBy = "li\"_"
                },
            };
            var csvFilePath = $@"{tempDirPath}\test.csv";
            entities.ToExcelFileByTemplate(
                Path.Combine(ApplicationHelper.AppRoot, "Templates", "testTemplate.xlsx"),
                ApplicationHelper.MapPath("templateTestEntities.xlsx"),
                extraData: new
                {
                    Author = "WeihanLi",
                    Title = "Export Result"
                }
            );
            entities.ToExcelFile(csvFilePath.Replace(".csv", ".xlsx"));
            entities.ToCsvFile(csvFilePath);
            var entitiesT0 = ExcelHelper.ToEntityList<TestEntity>(csvFilePath.Replace(".csv", ".xlsx"));

            var dataTable = entities.ToDataTable();
            dataTable.ToCsvFile(csvFilePath.Replace(".csv", ".datatable.csv"));
            var dt = CsvHelper.ToDataTable(csvFilePath.Replace(".csv", ".datatable.csv"));
            Console.WriteLine(dt.Columns.Count);
            var entities1 = CsvHelper.ToEntityList<TestEntity>(csvFilePath);

            entities1[1]!.DisplayName = ",tadadada";
            entities1[0]!.SettingValue = "value2,345";
            entities1.ToCsvFile(csvFilePath.Replace(".csv", ".1.csv"));
            entities1.ToDataTable().ToCsvFile(csvFilePath.Replace(".csv", ".1.datatable.csv"));

            var list = CsvHelper.ToEntityList<TestEntity>(csvFilePath.Replace(".csv", ".1.csv"));
            dt = CsvHelper.ToDataTable(csvFilePath.Replace(".csv", ".1.datatable.csv"));
            Console.WriteLine(dt.Columns.Count);
            var entities2 = CsvHelper.ToEntityList<TestEntity>(csvFilePath.Replace(".csv", ".1.csv"));

            entities.ToExcelFile(csvFilePath.Replace(".csv", ".xlsx"));

            var vals = new[] { 1, 2, 3, 5, 4 };
            vals.ToCsvFile(csvFilePath);

            var numList = CsvHelper.ToEntityList<int>(csvFilePath);
            Console.WriteLine(numList.StringJoin(","));

            Console.ReadLine();
        }

        private static void FluentSettingsForExcel()
        {
            var setting = FluentSettings.For<TestEntity>();
            // ExcelSetting
            setting.HasAuthor("WeihanLi")
                .HasTitle("WeihanLi.Npoi test")
                .HasDescription("WeihanLi.Npoi test")
                .HasSubject("WeihanLi.Npoi test");

            setting.HasSheetSetting(config =>
            {
                config.StartRowIndex = 1;
                config.SheetName = "SystemSettingsList";
                config.AutoColumnWidthEnabled = true;
            });

            // setting.HasFilter(0, 1).HasFreezePane(0, 1, 2, 1);

            setting.Property(_ => _.SettingId)
                .HasColumnIndex(0);

            setting.Property(_ => _.SettingName)
                .HasColumnTitle("SettingName")
                .HasColumnIndex(1);

            setting.Property(_ => _.DisplayName)
                .HasOutputFormatter((entity, displayName) => $"AAA_{entity?.SettingName}_{displayName}")
                .HasInputFormatter((entity, originVal) => originVal?.Split(new[] { '_' })[2])
                .HasColumnTitle("DisplayName")
                .HasColumnIndex(2);

            setting.Property(_ => _.SettingValue)
                .HasColumnTitle("SettingValue")
                .HasColumnIndex(3);

            setting.Property(_ => _.CreatedTime)
                .HasColumnTitle("CreatedTime")
                .HasColumnIndex(4)
                .HasColumnWidth(10)
                .HasColumnFormatter("yyyy-MM-dd HH:mm:ss");

            setting.Property(_ => _.CreatedBy)
                .HasColumnInputFormatter(x => x += "_test")
                .HasColumnIndex(4)
                .HasColumnTitle("CreatedBy");

            setting.Property(x => x.Enabled)
                .HasColumnInputFormatter(val => "Enabled".EqualsIgnoreCase(val))
                .HasColumnOutputFormatter(v => v ? "Enabled" : "Disabled");

            setting.Property("HiddenProp")
                .HasOutputFormatter((entity, val) => $"HiddenProp_{entity?.PKID}");

            setting.Property(_ => _.PKID).Ignored();
            setting.Property(_ => _.UpdatedBy).Ignored();
            setting.Property(_ => _.UpdatedTime).Ignored();
        }
    }

    internal abstract class BaseEntity
    {
        public int PKID { get; set; }
    }

    internal class TestEntity : BaseEntity
    {
        public Guid SettingId { get; set; }

        public string? SettingName { get; set; }

        public string? DisplayName { get; set; }
        public string? SettingValue { get; set; }

        public string CreatedBy { get; set; } = "liweihan";

        public DateTime CreatedTime { get; set; } = DateTime.Now;

        public string? UpdatedBy { get; set; }

        public DateTime UpdatedTime { get; set; }

        public bool Enabled { get; set; }
    }

    [Sheet(SheetIndex = 0, SheetName = "TestSheet", AutoColumnWidthEnabled = true)]
    internal class TestEntity2
    {
        [Column(Index = 0)]
        public int Id { get; set; }

        [Column(Index = 1)]
        public string? Title { get; set; }

        [Column(Index = 2, Width = 50)]
        public string? Description { get; set; }

        [Column(Index = 3, Width = 20)]
        public string? Extra { get; set; } = "{}";
    }
}
