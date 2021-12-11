using System.Collections.Generic;
using System.Reflection;
using Xunit.Sdk;

namespace WeihanLi.Npoi.Test;

public class ExcelFormatDataAttribute : DataAttribute
{
    public override IEnumerable<object[]> GetData(MethodInfo testMethod)
    {
        yield return new object[] { ExcelFormat.Xls };
        yield return new object[] { ExcelFormat.Xlsx };
    }
}
