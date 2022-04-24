// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System.Reflection;
using Xunit.Sdk;

namespace WeihanLi.Npoi.Test;

public sealed class ExcelFormatDataAttribute : DataAttribute
{
    public override IEnumerable<object[]> GetData(MethodInfo testMethod)
    {
        yield return new object[] { ExcelFormat.Xls };
        yield return new object[] { ExcelFormat.Xlsx };
    }
}
