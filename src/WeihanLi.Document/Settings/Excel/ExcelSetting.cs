namespace WeihanLi.Document.Settings.Excel
{
    /// <summary>
    /// Excel Settings
    /// Excel 文档属性设置
    /// </summary>
    public sealed class ExcelSetting
    {
        /// <summary>
        /// Author
        /// </summary>
        public string Author { get; set; } = "WeihanLi";

        /// <summary>
        /// Company
        /// </summary>
        public string Company { get; set; } = "WeihanLi";

        /// <summary>
        /// Title
        /// </summary>
        public string Title { get; set; } = "WeihanLi.Npoi";

        /// <summary>
        /// Description
        /// </summary>
        public string Description { get; set; } = "WeihanLi.Npoi";

        /// <summary>
        /// Subject
        /// </summary>
        public string Subject { get; set; } = "WeihanLi.Npoi";

        /// <summary>
        /// Category
        /// </summary>
        public string Category { get; set; } = "WeihanLi.Npoi";

        /// <summary>
        /// Default Excel Setting
        /// </summary>
        public static ExcelSetting Default = new ExcelSetting();
    }
}
