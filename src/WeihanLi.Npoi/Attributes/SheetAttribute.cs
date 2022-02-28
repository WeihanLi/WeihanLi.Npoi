// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the MIT license.

using System;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Attributes;

[AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
public sealed class SheetAttribute : Attribute
{
    private int _endColumnIndex = -1;

    private int _startColumnIndex;

    public SheetAttribute() => SheetSetting = new SheetSetting();
    public int SheetIndex { get; set; }

    public string SheetName
    {
        get => SheetSetting.SheetName;
        set => SheetSetting.SheetName = value ?? throw new ArgumentNullException(nameof(value));
    }

    public int StartRowIndex
    {
        get => SheetSetting.StartRowIndex;
        set => SheetSetting.StartRowIndex = value;
    }

    public int HeaderRowIndex => SheetSetting.HeaderRowIndex;

    public int EndRowIndex
    {
        get => SheetSetting.EndRowIndex ?? -1;
        set => SheetSetting.EndRowIndex = value >= 0 ? value : -1;
    }

    /// <summary>
    ///     StartColumnIndex
    ///     Start Column Index when import
    /// </summary>
    public int StartColumnIndex
    {
        get => _startColumnIndex;
        set
        {
            if (value >= 0)
            {
                _startColumnIndex = value;
                if (_endColumnIndex >= value)
                {
                    SheetSetting.CellFilter = cell =>
                            cell.ColumnIndex >= _startColumnIndex && cell.ColumnIndex <= _endColumnIndex
                        ;
                }
                else
                {
                    SheetSetting.CellFilter = cell => cell.ColumnIndex >= _startColumnIndex;
                }
            }
        }
    }

    /// <summary>
    ///     EndColumnIndex
    ///     End Column Index when import
    /// </summary>
    public int EndColumnIndex
    {
        get => _endColumnIndex;
        set
        {
            if (value >= _startColumnIndex)
            {
                _endColumnIndex = value;
                SheetSetting.CellFilter = cell =>
                        cell.ColumnIndex >= _startColumnIndex && cell.ColumnIndex <= _endColumnIndex
                    ;
            }
            else
            {
                SheetSetting.CellFilter = cell => cell.ColumnIndex >= _startColumnIndex;
                _endColumnIndex = -1;
            }
        }
    }

    public bool AutoColumnWidthEnabled
    {
        get => SheetSetting.AutoColumnWidthEnabled;
        set => SheetSetting.AutoColumnWidthEnabled = value;
    }

    internal SheetSetting SheetSetting { get; }
}
