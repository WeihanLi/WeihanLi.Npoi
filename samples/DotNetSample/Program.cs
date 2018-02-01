using System;
using System.IO;
using System.Linq;
using WeihanLi.Npoi;
using WeihanLi.Npoi.Attributes;

namespace DotNetSample
{
    internal class Program
    {
        private const string testDbConnString = "server=.;uid=liweihan;pwd=Admin888;database=AccountingApp";
        private const string FilePath = @"C:\Users\liweihan.TUHU\Desktop\temp\tempFiles\\所有单店.xlsx";

        private static void Main(string[] args)
        {
            //var conn = new SqlConnection(testDbConnString);
            //var entityList = conn.Select<TestEntity>("select * from Users");
            //entityList[0].Amount = 0;
            //entityList[0].PasswordHash = "";
            ////var dataTable = entityList.ToDataTable();
            //var result = ExcelHelper.ExportToExcel(ConfigurationHelper.MapPath("test.xlsx"), entityList);
            //var result1 = ExcelHelper.ToEntityList<TestEntity>(ConfigurationHelper.MapPath("test.xlsx"));
            //// 找不到文件
            ////var aaa = ExcelHelper.ToEntityList<TestEntity>("");
            //var entityList1 = conn.Select<TestEntity2>("select * from Bills");
            //var result2 = ExcelHelper.ExportToExcel(ConfigurationHelper.MapPath("test1.xls"), entityList1);
            //entityList1 = ExcelHelper.ToEntityList<TestEntity2>(ConfigurationHelper.MapPath("test1.xls"));
            //

            var entityList = ExcelHelper.ToEntityList<Model>(FilePath).Where(_ => !string.IsNullOrWhiteSpace(_.HotelId)).ToArray();
            if (entityList.Length > 0)
            {
                var dir = Path.GetDirectoryName(FilePath);
                foreach (var group in entityList.GroupBy(e => new
                {
                    HotelId = e.HotelId.Trim(),
                    HotelName = e.HotelName.Trim()
                }))
                {
                    var path = $"{dir}\\{group.Key.HotelName}.xlsx";
                    group.ToArray().ToExcelFile(path);
                }
                Console.WriteLine("Success");
            }
            Console.ReadLine();
        }
    }

    internal class TestEntity
    {
        /// <summary>
        /// 用户名
        /// </summary>
        [Column("用户名", Index = 0)]
        public string Username { get; set; }

        [Column("密码", Index = 1)]
        public string PasswordHash { get; set; }

        [Column("可用余额", Index = 2)]
        public decimal Amount { get; set; } = 1000M;

        [Column("微信id", Index = 3)]
        public string WechatOpenId { get; set; }

        [Column("是否启用", Index = 4)]
        public bool IsActive { get; set; }
    }

    internal class TestEntity2
    {
        [Column("PKID", Index = 0)]
        public int PKID { get; set; }

        [Column("账单标题", Index = 1)]
        public string BillTitle { get; set; }

        [Column("账单详情", Index = 2)]
        public string BillDetails { get; set; }

        [Column("创建人", Index = 3)]
        public string CreatedBy { get; set; }

        [Column("创建时间", Index = 4)]
        public DateTime CreatedTime { get; set; }
    }
}
