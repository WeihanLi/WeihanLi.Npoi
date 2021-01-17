namespace WeihanLi.Npoi.Abstract
{
    internal interface IRow
    {
        /// <summary>
        /// Gets the number of defined cells (NOT number of cells in the actual row!).
        /// That is to say if only columns 0,4,5 have values then there would be 3.
        /// </summary>
        /// <returns>int representing the number of defined cells in the row.</returns>
        int CellsCount { get; }

        /// <summary>
        ///   1-based column number of the first cell
        /// </summary>
        int FirstCellNum { get; }

        /// <summary>
        ///  1-based column number of the last cell
        /// </summary>
        int LastCellNum { get; }

        ICell? GetCell(int cellIndex);

        /// <summary>
        /// Create a cell
        /// </summary>
        /// <param name="cellIndex"> cellIndex
        /// maxValue: (255 for *.xls, 1048576 for *.xlsx)
        /// </param>
        /// <returns></returns>
        ICell CreateCell(int cellIndex);

        /// <summary>
        /// UnderlyingValue
        /// </summary>
        object? UnderlyingValue { get; }
    }
}
