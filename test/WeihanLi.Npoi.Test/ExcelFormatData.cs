// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using Xunit;

namespace WeihanLi.Npoi.Test;

public sealed class ExcelFormatData : TheoryData<ExcelFormat>
{
    public ExcelFormatData()
    {
        Add(ExcelFormat.Xls);
        Add(ExcelFormat.Xlsx);
    }
}
