namespace WeihanLi.Npoi.Settings
{
    /// <summary>
    /// Excel 设置
    /// SummaryInfo
    /// </summary>
    public sealed class ExcelSetting
    {
        public string Author { get; set; }

        public string Title { get; set; }

        public string Description { get; set; }

        public string Subject { get; set; }

        public ExcelSetting()
        {
            Author = "WeihanLi";
            Title = "WeihanLi.Npoi Generated";
            Description = "WeihanLi.Npoi Generated";
            Subject = "WeihanLi.Npoi Generated";
        }

        #region FluentAPI

        public ExcelSetting HasAuthor(string author)
        {
            Author = author;
            return this;
        }

        public ExcelSetting HasTitle(string title)
        {
            Title = title;
            return this;
        }

        public ExcelSetting HasDescription(string description)
        {
            Description = description;
            return this;
        }

        public ExcelSetting HasSubject(string subject)
        {
            Subject = subject;
            return this;
        }

        #endregion FluentAPI
    }
}
