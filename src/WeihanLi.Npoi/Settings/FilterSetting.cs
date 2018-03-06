namespace WeihanLi.Npoi.Settings
{
    internal sealed class FilterSetting
    {
        ///// <summary>
        ///// Gets or sets the first row index.
        ///// </summary>
        //public int FirstRow { get; }

        ///// <summary>
        ///// Gets or sets  the last row index.
        ///// </summary>
        ///// <remarks>
        ///// If the <see cref="LastRow"/> is null, the value is dynamic calculate by code.
        ///// </remarks>
        //public int? LastRow { get; }

        /// <summary>
        /// Gets or sets the first column index.
        /// </summary>
        public int FirstColumn { get; }

        /// <summary>
        /// Gets or sets the last column index.
        /// </summary>
        public int? LastColumn { get; }

        public FilterSetting(int firstColumn) : this(firstColumn, null)
        {
        }

        public FilterSetting(int firstColumn, int? lastColumn)
        {
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
        }
    }
}
