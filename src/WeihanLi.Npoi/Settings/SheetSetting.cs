// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using NPOI.SS.UserModel;
using WeihanLi.Extensions;

namespace WeihanLi.Npoi.Settings;

/// <summary>
///     Excel Sheet Settings
/// </summary>
public sealed class SheetSetting
{
    private Func<ICell, bool> _cellFilter = _ => true;

    /// <summary>
    ///     SheetName
    /// </summary>
    public string SheetName
    {
        get;
        set
        {
            if (value.IsNotNullOrWhiteSpace())
            {
                field = value;
            }
        }
    } = "Sheet0";

    /// <summary>
    ///     StartRowIndex
    /// </summary>
    public int StartRowIndex
    {
        get;
        set
        {
            if (value >= 0)
            {
                field = value;
            }
        }
    } = 1;

    /// <summary>
    ///     HeaderRowIndex
    /// </summary>
    public int HeaderRowIndex => StartRowIndex - 1;

    /// <summary>
    ///     EndRowIndex, included
    /// </summary>
    public int? EndRowIndex { get; set; }

    /// <summary>
    ///    Gets or set whether to enable auto column width, disabled by default.
    /// </summary>
    public bool AutoColumnWidthEnabled { get; set; }

    /// <summary>
    ///    Gets or sets whether to skip column index adjustment based on header row during import.
    ///    When false (default), column indices are automatically adjusted based on header row.
    ///    When true, column indices are used as-is.
    /// </summary>
    public bool SkipHeaderRow { get; set; }

    /// <summary>
    ///     Cell Filter
    /// </summary>
    public Func<ICell, bool>? CellFilter
    {
        get => _cellFilter;
        set => _cellFilter = value ?? (_ => true);
    }

    /// <summary>
    ///     Row Filter
    /// </summary>
    public Func<IRow, bool>? RowFilter
    {
        get;
        set => field = value ?? (_ => true);
    } = _ => true;

    /// <summary>
    ///     Cell Action on export
    /// </summary>
    public Action<ICell>? CellAction { get; set; }

    /// <summary>
    ///     Row Action on export
    /// </summary>
    public Action<IRow>? RowAction { get; set; }

    /// <summary>
    ///     Sheet Action on export
    /// </summary>
    public Action<ISheet>? SheetAction { get; set; }
}
