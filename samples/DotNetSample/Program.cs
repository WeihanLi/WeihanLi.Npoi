using System;
using System.Data.SqlClient;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;
using WeihanLi.Npoi;
using WeihanLi.Npoi.Attributes;

namespace DotNetSample
{
    internal class Program
    {
        private static string testDbConnString = "server=.;uid=liweihan;pwd=Admin888;database=AccountingApp";

        private static void Main(string[] args)
        {
            var conn = new SqlConnection(testDbConnString);
            var entityList = conn.Select<TestEntity>("select * from Users");
            var dataTable = entityList.ToDataTable();
            var result = ExcelHelper.ExportToExcel(ConfigurationHelper.MapPath("test.xlsx"), entityList);

            var result1 = ExcelHelper.ExportToExcel<TestEntity>(ConfigurationHelper.MapPath("test1.xlsx"), dataTable);
            // 找不到文件
            //var aaa = ExcelHelper.ToEntityList<TestEntity>("");
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

        /// <summary>
        /// 密码
        /// </summary>
        [Column("密码", Index = 1)]
        public string PasswordHash { get; set; }

        [Column("微信id", Index = 4)]
        public string WechatOpenId { get; set; }

        [Column("是否启用", Index = 3)]
        public bool IsActive { get; set; }
    }
}