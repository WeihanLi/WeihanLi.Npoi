using WeihanLi.Npoi.Test.MappingProfiles;

namespace WeihanLi.Npoi.Test
{
    public class Startup
    {
        public void Configure()
        {
            // ---------- load excel mapping profiles ----------------
            FluentSettings.LoadMappingProfiles(typeof(NoticeProfile).Assembly);
        }
    }
}
