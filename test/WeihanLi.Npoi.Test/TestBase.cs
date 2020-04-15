using WeihanLi.Extensions;
using WeihanLi.Npoi.Test.Models;
using Xunit;

namespace WeihanLi.Npoi.Test
{
    public class TestFixture
    {
        public TestFixture()
        {
            FluentSettingsConfigure();
        }

        private void FluentSettingsConfigure()
        {
            // ---------- notice npoi settings ----------------
            var noticeSetting = FluentSettings.For<Notice>();
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

    [CollectionDefinition("Tests")]
    public class NoneParallelTestCollection : ICollectionFixture<TestFixture>
    {
    }

    [Collection("Tests")]
    public class TestBase
    {
    }
}
