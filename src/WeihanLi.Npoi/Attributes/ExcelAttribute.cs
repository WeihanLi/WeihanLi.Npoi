using System;

namespace WeihanLi.Npoi.Attributes
{
    /// <summary>
    /// Excel 信息
    /// SummaryInfo
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public sealed class ExcelAttribute : Attribute
    {
        public string Author { get; set; }
    }
}