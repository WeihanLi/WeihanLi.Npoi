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
        internal int ColumnIndex { get; set; }

        /// <summary>
        /// Title
        /// </summary>
        internal string ColumnTitle { get; set; }

        /// <summary>
        /// Formatter
        /// </summary>
        internal string ColumnFormatter { get; set; }

        /// <summary>
        /// the property is ignored.
        /// </summary>
        internal bool IsIgnored { get; set; }
    }
}
