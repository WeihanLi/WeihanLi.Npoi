namespace WeihanLi.Npoi.Settings
{
    /// <summary>
    /// PropertySetting
    /// </summary>
    internal class PropertySetting
    {
        /// <summary>
        /// ColumnIndex
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// Title
        /// </summary>
        public string ColumnTitle { get; set; }

        /// <summary>
        /// Formatter
        /// </summary>
        public string ColumnFormatter { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this value of the property is ignored.
        /// </summary>
        /// <value><c>true</c> if is ignored; otherwise, <c>false</c>.</value>
        public bool IsIgnored { get; set; }
    }
}
