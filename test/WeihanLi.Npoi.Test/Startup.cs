// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using WeihanLi.Npoi.Test.MappingProfiles;

namespace WeihanLi.Npoi.Test;

public class Startup
{
    public void Configure()
    {
        System.AppContext.SetSwitch("System.Drawing.EnableUnixSupport", true);
        // ---------- load excel mapping profiles ----------------
        FluentSettings.LoadMappingProfiles(typeof(NoticeProfile).Assembly);
    }
}
