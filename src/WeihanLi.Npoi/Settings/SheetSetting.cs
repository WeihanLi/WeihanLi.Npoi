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
    private Func<IRow, bool>? _rowFilter = _ => true;
    private string _sheetName = "Sheet0";
    private int _startRowIndex = 1;

    /// <summary>
    ///     SheetName
    /// </summary>
    public string SheetName
    {
        get => _sheetName;
        set
        {
            if (value.IsNotNullOrWhiteSpace())
            {
                _sheetName = value;
            }
        }
    }

    /// <summary>
    ///     StartRowIndex
    /// </summary>
    public int StartRowIndex
    {
        get => _startRowIndex;
        set
        {
            if (value >= 0)
            {
                _startRowIndex = value;
            }
        }
    }

    /// <summary>
    ///     HeaderRowIndex
    /// </summary>
    public int HeaderRowIndex => StartRowIndex - 1;

    /// <summary>
    ///     EndRowIndex, included
    /// </summary>
    public int? EndRowIndex { get; set; }

    /// <summary>
    ///     enable auto column width
    /// </summary>
    public bool AutoColumnWidthEnabled { get; set; }

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
        get => _rowFilter;
        set => _rowFilter = value ?? (_ => true);
    }

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
