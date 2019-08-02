namespace WeihanLi.Npoi.Settings
{
    internal sealed class FilterSetting
    {
        /// <summary>
        /// Gets or sets the first column index.
        /// </summary>
        public int FirstColumn { get; }

        /// <summary>
        /// Gets or sets the last column index.
        /// </summary>
        public int? LastColumn { get; }

        public FilterSetting(int firstColumn, int? lastColumn)
        {
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
        }
    }
}
