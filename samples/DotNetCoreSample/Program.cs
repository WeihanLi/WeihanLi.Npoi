using System;
using System.Data.SqlClient;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;
using WeihanLi.Npoi;

// ReSharper disable All
namespace DotNetCoreSample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            FluentSettingsForExcel();

            using (var conn = new SqlConnection("server=.;uid=liweihan;pwd=Admin888;database=Reservation"))
            {
                var list = conn.Select<TestEntity>(@"SELECT * FROM [Reservation].[dbo].[tabSystemSettings]");
                list.ToExcelFile(ConfigurationHelper.MapPath("test.xlsx"));
            }

            Console.WriteLine("Success!");
            Console.ReadLine();
        }

        private static void FluentSettingsForExcel()
        {
            var setting = ExcelHelper.SettingFor<TestEntity>();
            // ExcelSetting
            setting.HasAuthor("WeihanLi")
                .HasTitle("WeihanLi.Npoi test")
                .HasDescription("")
                .HasSubject("");

            setting.HasFilter(0, 1)
                .HasFreezePane(0, 1, 2, 1);
            setting.Property(_ => _.SettingId)
                .HasColumnIndex(0);

            setting.Property(_ => _.SettingName)
                .HasColumnTitle("设置名称")
                .HasColumnIndex(1);

            setting.Property(_ => _.DisplayName)
                .HasColumnTitle("设置显示名称")
                .HasColumnIndex(2);

            setting.Property(_ => _.SettingValue)
                .HasColumnTitle("设置值")
                .HasColumnIndex(3);

            setting.Property(_ => _.CreatedTime)
                .HasColumnTitle("创建时间")
                .HasColumnIndex(5)
                .HasColumnFormatter("yyyy-MM-dd HH:mm:ss");

            setting.Property(_ => _.CreatedBy)
                .HasColumnIndex(4)
                .HasColumnTitle("创建人");

            setting.Property(_ => _.UpdatedBy)
                .Ignored();
            setting.Property(_ => _.UpdatedTime)
                .Ignored();
        }
    }

    internal class TestEntity
    {
        public Guid SettingId { get; set; }

        public string SettingName { get; set; }

        public string DisplayName { get; set; }
        public string SettingValue { get; set; }

        public string CreatedBy { get; set; } = "liweihan";

        public DateTime CreatedTime { get; set; } = DateTime.Now;

        public string UpdatedBy { get; set; }

        public DateTime UpdatedTime { get; set; }
    }
}
