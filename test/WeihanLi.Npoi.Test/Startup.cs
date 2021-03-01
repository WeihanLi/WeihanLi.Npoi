using WeihanLi.Extensions;
using WeihanLi.Npoi.Test.Models;

namespace WeihanLi.Npoi.Test
{
    public class Startup
    {
        public void Configure()
        {
            // ---------- notice fluent excel settings ----------------
            var noticeSetting = FluentSettings.For<Notice>();
            noticeSetting
                .HasAuthor("WeihanLi")
                .HasTitle("WeihanLi.Npoi test")
                .HasSheetSetting(setting =>
                {
                    setting.SheetName = "NoticeList";
                    setting.AutoColumnWidthEnabled = true;
                })
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
                .HasColumnOutputFormatter(x => x.ToStandardTimeString());
        }
    }
}
