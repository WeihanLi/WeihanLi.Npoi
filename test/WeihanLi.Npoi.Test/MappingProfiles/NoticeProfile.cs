// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using WeihanLi.Extensions;
using WeihanLi.Npoi.Configurations;
using WeihanLi.Npoi.Test.Models;

namespace WeihanLi.Npoi.Test.MappingProfiles;

public sealed class NoticeProfile : IMappingProfile<Notice>
{
    public void Configure(IExcelConfiguration<Notice> noticeSetting)
    {
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
            .HasColumnOutputFormatter(x => x.ToTimeString());
    }
}
