using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using EPPlus.Core.Extensions;
using EPPlus.Core.Extensions.Attributes;
using WeihanLi.Common.Helpers;
using WeihanLi.Npoi;
using WeihanLi.Npoi.Attributes;

// ReSharper disable LocalizableElement
namespace DotNetSample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //var conn = new SqlConnection("server=.;uid=liweihan;pwd=Admin888;database=AccountingApp");
            //var entityList = conn.Select<TestEntity>("select * from Users").ToArray();
            //entityList[0].Amount = 0;
            //entityList[0].PasswordHash = "";
            ////var dataTable = entityList.ToDataTable();
            //var result = entityList.ToExcelFile(ApplicationHelper.MapPath("test.xlsx"));
            //var result1 = ExcelHelper.ToEntityList<TestEntity>(ApplicationHelper.MapPath("test.xlsx"));
            //// 找不到文件
            ////var aaa = ExcelHelper.ToEntityList<TestEntity>("");
            //var entityList1 = conn.Select<TestEntity2>("select * from Bills");
            //var result2 = entityList1.ToExcelFile(ConfigurationHelper.MapPath("test1.xls"));
            //entityList1 = ExcelHelper.ToEntityList<TestEntity2>(ConfigurationHelper.MapPath("test1.xls"));

            //var entityList2 = ExcelHelper.ToEntityList<Model>(FilePath).Where(_ => !string.IsNullOrWhiteSpace(_.HotelId)).ToArray();
            //if (entityList2.Length > 0)
            //{
            //    var dir = Path.GetDirectoryName(FilePath);
            //    foreach (var group in entityList2.GroupBy(e => new
            //    {
            //        HotelId = e.HotelId.Trim(),
            //        HotelName = e.HotelName.Trim()
            //    }))
            //    {
            //        var path = $"{dir}\\sub\\{group.Key.HotelName}-1月对账单.xlsx";
            //        group.ToArray().ToExcelFile(path);
            //    }
            //    Console.WriteLine("Success");
            //}

            //var table = ExcelHelper.ToDataTable(FilePath);
            ////Console.WriteLine(table.Rows.Count);

            ////Console.WriteLine(table.ToExcelFile(FilePath.Replace("AllStores", "AllStores1")));

            //using (var connection = new SqlConnection("server=.;uid=liweihan;pwd=Admin888;database=TestDb"))
            //{
            //    Console.WriteLine($"导入结果：{connection.BulkCopy(table, "testBulkCopy")}");
            //}

            //var setting = FluentSettings.For<TestEntity>();
            //// ExcelSetting
            //setting.HasAuthor("WeihanLi")
            //    .HasTitle("WeihanLi.Npoi test")
            //    .HasDescription("")
            //    .HasSubject("");

            //setting.HasFilter(0, 1)
            //    .HasFreezePane(0, 1, 2, 1);

            //setting.Property(_ => _.Amount)
            //    .HasColumnTitle("Amount")
            //    .HasColumnIndex(2);

            //setting.Property(_ => _.Username)
            //    .HasColumnTitle("Username")
            //    .HasColumnIndex(0);

            //setting.Property(_ => _.CreateTime)
            //    .HasColumnTitle("CreateTime")
            //    .HasColumnFormatter("yyyy-MM-dd HH:mm:ss");

            //setting.Property(_ => _.PasswordHash)
            //    .Ignored();

            //var entities = ExcelHelper.ToEntityList<TestEntity>(ApplicationHelper.MapPath("test.xlsx"));
            //Console.WriteLine(entities.Count);
            ////entities = conn.Select<TestEntity>("select * from Users").ToList();
            //entities.ToExcelFile(ApplicationHelper.MapPath("test_1.xlsx"));
            //Console.WriteLine("Success");

            //Console.WriteLine($"WorkingSet size: {Process.GetCurrentProcess().WorkingSet64 / 1024} kb");

            //// ExportExcelViaEpplusPerfTest();
            //// ExportExcelViaEpplusPerfTest(1_000_000, 5);

            //// ExportExcelPerfTest(100_000, 10);
            //// ExportExcelPerfTest(1_000_000, 5);

            //ExportCsvPerfTest(100_000, 10);
            //// ExportCsvPerfTest(1_000_000, 5);

            //Console.WriteLine($"WorkingSet size: {Process.GetCurrentProcess().WorkingSet64 / 1024} kb");
            //GC.Collect(2, GCCollectionMode.Forced);
            //Console.WriteLine($"WorkingSet size: {Process.GetCurrentProcess().WorkingSet64 / 1024} kb");

            var testData = new List<TestEntity>(10);

            for (int i = 1; i <= 10; i++)
            {
                testData.Add(new TestEntity()
                {
                    Amount = 1000,
                    Username = "xxxx",
                    CreateTime = DateTime.UtcNow.AddDays(-3),
                    PKID = i,
                    PasswordHash = SecurityHelper.SHA1($"_x_{i}")
                });
            }
            var excelFilePath = $@"{Environment.GetEnvironmentVariable("USERPROFILE")}\Desktop\temp\test\test123.fx.xlsx";
            testData.ToExcelFile(excelFilePath);
            testData.ToCsvFile(excelFilePath.Replace(".xlsx", ".csv"));

            var list = ExcelHelper.ToEntityList<TestEntity>(excelFilePath);

            Console.WriteLine("complete");
            Console.ReadLine();
        }

        private static void ExportExcelPerfTest(int recordCount, int repeatTimes = 10)
        {
            if (recordCount <= 0)
                recordCount = 100_000;

            var testData = new List<TestEntity>(recordCount);

            for (int i = 1; i <= recordCount; i++)
            {
                testData.Add(new TestEntity()
                {
                    Amount = 1000,
                    Username = "xxxx",
                    CreateTime = DateTime.UtcNow.AddDays(-3),
                    PKID = i,
                });
            }
            var stopwatch = new Stopwatch();
            var excelFilePath = $@"{Environment.GetEnvironmentVariable("USERPROFILE")}\Desktop\temp\test\test.fx.xlsx";
            var elapsedList = new List<long>(repeatTimes);
            for (int i = 0; i < repeatTimes; i++)
            {
                stopwatch.Restart();
                testData.ToExcelFile(excelFilePath);
                stopwatch.Stop();
                elapsedList.Add(stopwatch.ElapsedMilliseconds);
                Console.WriteLine($"{stopwatch.ElapsedMilliseconds} ms");
            }
            Console.WriteLine($"Average: {elapsedList.Average()} ms");
        }

        private static void ExportExcelViaEpplusPerfTest(int recordCount = -1, int repeatTimes = 10)
        {
            if (recordCount <= 0)
                recordCount = 100_000;

            var testData = new List<TestEntity>(recordCount);

            for (int i = 1; i <= recordCount; i++)
            {
                testData.Add(new TestEntity()
                {
                    Amount = 1000,
                    Username = "xxxx",
                    CreateTime = DateTime.UtcNow.AddDays(-3),
                    PKID = i,
                });
            }
            var stopwatch = new Stopwatch();
            var excelFilePath = $@"{Environment.GetEnvironmentVariable("USERPROFILE")}\Desktop\temp\test\test.fx.epplus.xlsx";
            var elapsedList = new List<long>(repeatTimes);
            for (int i = 0; i < repeatTimes; i++)
            {
                stopwatch.Restart();
                testData.ToExcelPackage().SaveAs(new System.IO.FileInfo(excelFilePath));
                stopwatch.Stop();
                elapsedList.Add(stopwatch.ElapsedMilliseconds);
                Console.WriteLine($"{stopwatch.ElapsedMilliseconds} ms");
            }
            Console.WriteLine($"Average: {elapsedList.Average()} ms");
        }

        private static void ExportCsvPerfTest(int recordCount, int repeatTimes = 10)
        {
            if (recordCount <= 0)
                recordCount = 100_000;

            var testData = new List<TestEntity>(recordCount);

            for (int i = 1; i <= recordCount; i++)
            {
                testData.Add(new TestEntity()
                {
                    Amount = 1000,
                    Username = "xxxx",
                    CreateTime = DateTime.UtcNow.AddDays(-3),
                    PKID = i,
                });
            }
            var stopwatch = new Stopwatch();
            var excelFilePath = $@"{Environment.GetEnvironmentVariable("USERPROFILE")}\Desktop\temp\test\test.fx.csv";
            var elapsedList = new List<long>(repeatTimes);
            for (int i = 0; i < repeatTimes; i++)
            {
                stopwatch.Restart();
                testData.ToCsvFile(excelFilePath);
                stopwatch.Stop();
                elapsedList.Add(stopwatch.ElapsedMilliseconds);
                Console.WriteLine($"{stopwatch.ElapsedMilliseconds} ms");
            }
            Console.WriteLine($"Average: {elapsedList.Average()} ms");
        }
    }

    //[Freeze(0, 1)]
    //[Sheet(SheetIndex = 0, SheetName = "Abc", StartRowIndex = 0)]
    internal class TestEntity
    {
        [ExcelTableColumn("PKID")]
        public int PKID { get; set; }

        /// <summary>
        /// 用户名
        /// </summary>
        [Column("Username")]
        [ExcelTableColumn("UserName")]
        public string Username { get; set; }

        [Column("PasswordHash")]
        [ExcelTableColumn("PasswordHash")]
        public string PasswordHash { get; set; }

        [Column("Amount")]
        [ExcelTableColumn("Amount")]
        public decimal Amount { get; set; } = 1000M;

        [Column("WechatOpenId")]
        [ExcelTableColumn("WechatOpenId")]
        public string WechatOpenId { get; set; }

        [Column("IsActive")]
        [ExcelTableColumn("IsActive")]
        public bool IsActive { get; set; }

        [ExcelTableColumn("CreateTime")]
        public DateTime CreateTime { get; set; } = DateTime.Now;
    }

    internal class TestEntity2
    {
        [Column("ID")]
        public int PKID { get; set; }

        [Column("BillTitle")]
        public string BillTitle { get; set; }

        [Column("BillDetails")]
        public string BillDetails { get; set; }

        [Column("CreatedBy")]
        public string CreatedBy { get; set; }

        [Column("CreatedTime", Formatter = "yyyy-MM-dd HH:mm:ss")]
        public DateTime CreatedTime { get; set; }
    }
}
