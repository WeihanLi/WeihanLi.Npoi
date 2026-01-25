// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using NPOI.SS.UserModel;
using System.Collections;
using WeihanLi.Common;

namespace WeihanLi.Npoi;

/// <summary>
///     npoi sheet row collection
/// </summary>
public sealed class NpoiRowCollection(ISheet sheet) : IReadOnlyCollection<IRow>
{
    private readonly ISheet _sheet = Guard.NotNull(sheet);

    /// <summary>
    ///     Gets the number of rows within the wrapped sheet.
    /// </summary>
    public int Count => _sheet.LastRowNum - _sheet.FirstRowNum + 1;

    /// <summary>
    ///     Iterates over each row within the sheet range.
    /// </summary>
    public IEnumerator<IRow> GetEnumerator()
    {
        for (var i = _sheet.FirstRowNum; i <= _sheet.LastRowNum; i++)
        {
            yield return _sheet.GetRow(i);
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}

/// <summary>
///     npoi row cell collection
/// </summary>
public sealed class NpoiCellCollection(IRow row) : IReadOnlyCollection<ICell>
{
    private readonly IRow _row = Guard.NotNull(row);

    /// <summary>
    ///     Gets the number of cells in the wrapped row.
    /// </summary>
    public int Count => _row.LastCellNum - _row.FirstCellNum;

    /// <summary>
    ///     Iterates over each concrete cell in the current row.
    /// </summary>
    public IEnumerator<ICell> GetEnumerator()
    {
        for (var i = _row.FirstCellNum; i < _row.LastCellNum; i++)
        {
            yield return _row.GetCell(i);
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
