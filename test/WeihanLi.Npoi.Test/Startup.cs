using WeihanLi.Extensions;
using WeihanLi.Npoi.Test.MappingProfiles;
using WeihanLi.Npoi.Test.Models;

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
