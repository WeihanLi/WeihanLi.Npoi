// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using NPOI.SS.Util;
using NModel = NPOI.SS.UserModel;

namespace WeihanLi.Npoi.Abstract;

/// <summary>
/// Thin adapter that exposes the internal NPOI workbook via the abstraction interfaces.
/// </summary>
internal class NPOIWorkbook : IWorkbook
{
    private readonly NModel.IWorkbook _workbook;

    /// <summary>
    ///     Creates a new adapter for the provided NPOI workbook.
    /// </summary>
    /// <param name="workbook">Underlying workbook instance.</param>
    public NPOIWorkbook(NModel.IWorkbook workbook) => _workbook = workbook;

    /// <summary>
    ///     Gets the number of sheets available.
    /// </summary>
    public int SheetCount => _workbook.NumberOfSheets;

    /// <summary>
    ///     Wraps the specified sheet index in an <see cref="ISheet" /> adapter.
    /// </summary>
    public ISheet GetSheet(int sheetIndex) => new NPOISheet(_workbook.GetSheetAt(sheetIndex));

    /// <summary>
    ///     Creates a new sheet and returns the adapter around it.
    /// </summary>
    public ISheet CreateSheet(string sheetName) => new NPOISheet(_workbook.CreateSheet(sheetName));

    /// <summary>
    ///     Serializes the workbook to bytes using helper extensions.
    /// </summary>
    public byte[] ToBytes() => _workbook.ToExcelBytes();
}

/// <summary>
/// Adapter for <see cref="NModel.ISheet" />.
/// </summary>
internal class NPOISheet : ISheet
{
    private readonly NModel.ISheet _sheet;

    /// <summary>
    ///     Initializes the adapter with the underlying sheet.
    /// </summary>
    public NPOISheet(NModel.ISheet sheet) => _sheet = sheet;

    /// <summary>
    ///     Gets the first row index using one-based indexing to align with the abstractions.
    /// </summary>
    public int FirstRowNum => _sheet.FirstRowNum + 1;

    /// <summary>
    ///     Gets the last row index using one-based indexing to align with the abstractions.
    /// </summary>
    public int LastRowNum => _sheet.LastRowNum + 1;

    /// <summary>
    ///     Retrieves the requested row and wraps it, if present.
    /// </summary>
    public IRow? GetRow(int rowIndex)
    {
        var nRow = _sheet.GetRow(rowIndex);
        if (null == nRow)
        {
            return null;
        }

        return new NPOIRow(nRow);
    }

    /// <summary>
    ///     Creates a new row and returns its adapter.
    /// </summary>
    public IRow CreateRow(int rowIndex) => new NPOIRow(_sheet.CreateRow(rowIndex));

    /// <summary>
    ///     Sets the column width in the underlying sheet.
    /// </summary>
    public void SetColumnWidth(int columnIndex, int width) => _sheet.SetColumnWidth(columnIndex, width);

    /// <summary>
    ///     Auto sizes the requested column.
    /// </summary>
    public void AutoSizeColumn(int columnIndex) => _sheet.AutoSizeColumn(columnIndex);

    /// <summary>
    ///     Applies a freeze pane to the sheet.
    /// </summary>
    public void CreateFreezePane(int colSplit, int rowSplit, int leftMostCol, int topRow) =>
        _sheet.CreateFreezePane(colSplit, rowSplit, leftMostCol, topRow);

    /// <summary>
    ///     Applies an auto-filter range to the sheet.
    /// </summary>
    public void SetAutoFilter(int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex) =>
        _sheet.SetAutoFilter(new CellRangeAddress(firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex));

    /// <summary>
    ///     Shifts the specified row range.
    /// </summary>
    public void ShiftRows(int startRow, int endRow, int n) => _sheet.ShiftRows(startRow, endRow, n);

    /// <summary>
    ///     Copies a row and wraps the result.
    /// </summary>
    public IRow CopyRow(int sourceIndex, int targetIndex) => new NPOIRow(_sheet.CopyRow(sourceIndex, targetIndex));

    /// <summary>
    ///     Removes the given row from the sheet.
    /// </summary>
    public void RemoveRow(IRow row) => _sheet.RemoveRow(row.UnderlyingValue as NModel.IRow);
}

/// <summary>
/// Adapter for <see cref="NModel.IRow" />.
/// </summary>
internal class NPOIRow : IRow
{
    private readonly NModel.IRow _row;

    /// <summary>
    ///     Initializes the adapter with the underlying row.
    /// </summary>
    public NPOIRow(NModel.IRow row) => _row = row;

    /// <summary>
    ///     Gets the number of physical cells.
    /// </summary>
    public int CellsCount => _row.PhysicalNumberOfCells;

    /// <summary>
    ///     Gets the first cell index using one-based indexing.
    /// </summary>
    public int FirstCellNum => _row.FirstCellNum + 1;

    /// <summary>
    ///     Gets the last cell index.
    /// </summary>
    public int LastCellNum => _row.LastCellNum;

    /// <summary>
    ///     Retrieves and wraps the specified cell.
    /// </summary>
    public ICell? GetCell(int cellIndex)
    {
        var nCell = _row.GetCell(cellIndex);
        if (nCell is null)
        {
            return null;
        }

        return new NPOICell(nCell);
    }

    /// <summary>
    ///     Creates and wraps a new cell.
    /// </summary>
    public ICell CreateCell(int cellIndex) => new NPOICell(_row.CreateCell(cellIndex));

    /// <summary>
    ///     Provides direct access to the underlying NPOI object.
    /// </summary>
    public object UnderlyingValue => _row;
}

/// <summary>
/// Adapter for <see cref="NModel.ICell" />.
/// </summary>
internal class NPOICell : ICell
{
    private readonly NModel.ICell _cell;

    /// <summary>
    ///     Initializes the adapter with the underlying cell.
    /// </summary>
    public NPOICell(NModel.ICell cell) => _cell = cell;

    /// <summary>
    ///     Gets or sets the cell type using the abstraction enum.
    /// </summary>
    public CellType CellType
    {
        get => (CellType)Enum.Parse(typeof(CellType), _cell.CellType.ToString());
        set => _cell.SetCellType((NModel.CellType)Enum.Parse(typeof(NModel.CellType), value.ToString()));
    }

    /// <summary>
    ///     Gets or sets the cell value while translating common NPOI types.
    /// </summary>
    public object? Value
    {
        get
        {
            if (_cell.CellType == NModel.CellType.Blank || _cell.CellType == NModel.CellType.Error)
            {
                return null;
            }

            switch (_cell.CellType)
            {
                case NModel.CellType.Numeric:
                    if (NModel.DateUtil.IsCellDateFormatted(_cell))
                    {
                        return _cell.DateCellValue;
                    }

                    return _cell.NumericCellValue;

                case NModel.CellType.String:
                    return _cell.StringCellValue;

                case NModel.CellType.Boolean:
                    return _cell.BooleanCellValue;

                default:
                    return _cell.ToString();
            }
        }

        set => _cell.SetCellValue(value ?? string.Empty);
    }
}
