// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi;

public interface IMappingProfile
{
}

public interface IMappingProfile<T> : IMappingProfile
{
    public void Configure(IExcelConfiguration<T> configuration);
}
