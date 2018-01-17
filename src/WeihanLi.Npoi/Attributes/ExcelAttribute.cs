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

        public string Title { get; set; }

        public string Description { get; set; }

        public string Subject { get; set; }

        public ExcelAttribute()
        {
            Author = "WeihanLi";
            Title = "WeihanLi.Npoi Generated";
            Description = "WeihanLi.Npoi Generated";
            Subject = "WeihanLi.Npoi Generated";
        }
    }
}
