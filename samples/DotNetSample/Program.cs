using System;
using System.Collections.Generic;
using System.Diagnostics;
using WeihanLi.Common.Helpers;
using WeihanLi.Npoi;
using WeihanLi.Npoi.Attributes;

// ReSharper disable LocalizableElement
namespace DotNetSample
{
    public class Program
    {
        private const string FilePath = @"C:\Users\liweihan.TUHU\Desktop\temp\tempFiles\\AllStores.xlsx";

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

            var testData = new List<TestEntity>()
            {
                new TestEntity
                {
                    Amount = 1000,
                    Username = "xxxx",
                    CreateTime = DateTime.UtcNow.AddDays(-3),
                    PKID = 1,
                },
                new TestEntity
                {
                    Amount = 10000,
                    Username = "yyyyy",
                    CreateTime = DateTime.UtcNow.AddDays(-3),
                    PKID = 2,
                }
            };
            for (int i = 0; i < 100000; i++)
            {
                testData.Add(new TestEntity()
                {
                    Amount = 1000,
                    Username = "xxxx",
                    CreateTime = DateTime.UtcNow.AddDays(-3),
                    PKID = i + 2,
                });
            }
            var stopwatch = new Stopwatch();
            var excelFilePath = $@"{Environment.GetEnvironmentVariable("USERPROFILE")}\Desktop\temp\test\test.fx.xlsx";
            stopwatch.Start();
            testData.ToExcelFile(excelFilePath);
            stopwatch.Stop();
            Console.WriteLine(stopwatch.ElapsedMilliseconds);
            testData = null;

            var setting = ExcelHelper.SettingFor<TestEntity>();
            // ExcelSetting
            setting.HasAuthor("WeihanLi")
                .HasTitle("WeihanLi.Npoi test")
                .HasDescription("")
                .HasSubject("");

            setting.HasFilter(0, 1)
                .HasFreezePane(0, 1, 2, 1);

            setting.Property(_ => _.Amount)
                .HasColumnTitle("Amount")
                .HasColumnIndex(2);

            setting.Property(_ => _.Username)
                .HasColumnTitle("Username")
                .HasColumnIndex(0);

            setting.Property(_ => _.CreateTime)
                .HasColumnTitle("CreateTime")
                .HasColumnFormatter("yyyy-MM-dd HH:mm:ss");

            setting.Property(_ => _.PasswordHash)
                .Ignored();

            var entities = ExcelHelper.ToEntityList<TestEntity>(ApplicationHelper.MapPath("test.xlsx"));
            Console.WriteLine(entities.Count);
            //entities = conn.Select<TestEntity>("select * from Users").ToList();
            entities.ToExcelFile(ApplicationHelper.MapPath("test_1.xlsx"));
            Console.WriteLine("Success");

            Console.ReadLine();
        }
    }

    [Freeze(0, 1)]
    //[Sheet(SheetIndex = 0, SheetName = "Abc", StartRowIndex = 0)]
    internal class TestEntity
    {
        public int PKID { get; set; }

        /// <summary>
        /// 用户名
        /// </summary>
        [Column("Username")]
        public string Username { get; set; }

        [Column("PasswordHash", IsIgnored = true)]
        public string PasswordHash { get; set; }

        [Column("Amount")]
        public decimal Amount { get; set; } = 1000M;

        [Column("WechatOpenId")]
        public string WechatOpenId { get; set; }

        [Column("IsActive")]
        public bool IsActive { get; set; }

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
