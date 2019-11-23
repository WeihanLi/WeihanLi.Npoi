using WeihanLi.Extensions;
using WeihanLi.Npoi.Test.Models;

namespace WeihanLi.Npoi.Test
{
    public class TestBase
    {
        static TestBase()
        {
            // 初始化配置
            FluentSettings();
        }

        private static void FluentSettings()
        {
            // ---------- notice npoi settings ----------------
            var noticeSetting = ExcelHelper.SettingFor<Notice>();
            noticeSetting.HasAuthor("WeihanLi")
                .HasTitle("WeihanLi.Npoi test")
                .HasSheetConfiguration(0, "NoticeList", 1)
                ;
            noticeSetting.Property(_ => _.Id)
                .HasColumnIndex(0);
            noticeSetting.Property(_ => _.Title)
                .HasColumnIndex(1);
            noticeSetting.Property(_ => _.Content)
                .HasColumnIndex(2);
            noticeSetting.Property(_ => _.Publisher)
                .HasColumnIndex(3);
            noticeSetting.Property(_ => _.PublishedAt)
                .HasColumnIndex(4)
                .HasOutputFormatter((entity, x) => x.ToStandardTimeString());
        }
    }
}
