namespace WeihanLi.Npoi.Settings
{
    internal sealed class FilterSetting
    {
        public FilterSetting(int firstColumn, int? lastColumn)
        {
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
        }

        /// <summary>
        ///     Gets or sets the first column index.
        /// </summary>
        public int FirstColumn { get; }

        /// <summary>
        ///     Gets or sets the last column index.
        /// </summary>
        public int? LastColumn { get; }
    }
}
