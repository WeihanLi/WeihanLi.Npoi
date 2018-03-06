using System;
using System.Data.SqlClient;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;
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
            var conn = new SqlConnection("server=.;uid=liweihan;pwd=Admin888;database=AccountingApp");
            var entityList = conn.Select<TestEntity>("select * from Users");
            entityList[0].Amount = 0;
            entityList[0].PasswordHash = "";
            //var dataTable = entityList.ToDataTable();
            var result = entityList.ToExcelFile(ConfigurationHelper.MapPath("test.xlsx"));
            //var result1 = ExcelHelper.ToEntityList<TestEntity>(ConfigurationHelper.MapPath("test.xlsx"));
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

            var setting = ExcelHelper.SettingFor<TestEntity>();
            // ExcelSetting
            setting.HasAuthor("WeihanLi")
                .HasTitle("WeihanLi.Npoi")
                .HasDescription("")
                .HasSubject("");

            setting.HasFilter(0, 1)
                .HasFreezePane(0, 1);

            setting.Property(_ => _.Amount)
                .HasColumnTitle("可用余额")
                .HasColumnIndex(2);

            setting.Property(_ => _.Username)
                .HasColumnTitle("用户名")
                .HasColumnIndex(0);

            setting.Property(_ => _.CreateTime)
                .HasColumnTitle("创建时间")
                .HasColumnFormatter("yyyy-MM-dd HH:mm:ss");

            setting.Property(_ => _.PasswordHash)
                .Ignored();

            var entities = ExcelHelper.ToEntityList<TestEntity>(ConfigurationHelper.MapPath("test.xlsx"));
            Console.WriteLine(entities.Count);
            entities = conn.Select<TestEntity>("select * from Users");
            entities.ToExcelFile(ConfigurationHelper.MapPath("test_1.xlsx"));
            Console.WriteLine("Success");

            Console.ReadLine();
        }
    }

    [Freeze(0, 1)]
    internal class TestEntity
    {
        public int PKID { get; set; }

        /// <summary>
        /// 用户名
        /// </summary>
        [Column("用户名")]
        public string Username { get; set; }

        [Column("密码", IsIgnored = true)]
        public string PasswordHash { get; set; }

        [Column("可用余额")]
        public decimal Amount { get; set; } = 1000M;

        [Column("微信id")]
        public string WechatOpenId { get; set; }

        [Column("是否启用")]
        public bool IsActive { get; set; }

        public DateTime CreateTime { get; set; } = DateTime.Now;
    }

    internal class TestEntity2
    {
        [Column("PKID")]
        public int PKID { get; set; }

        [Column("账单标题")]
        public string BillTitle { get; set; }

        [Column("账单详情")]
        public string BillDetails { get; set; }

        [Column("创建人")]
        public string CreatedBy { get; set; }

        [Column("创建时间", Formatter = "yyyy-MM-dd HH:mm:ss")]
        public DateTime CreatedTime { get; set; }
    }
}
