// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Attributes;

[AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
public sealed class FreezeAttribute : Attribute
{
    public FreezeAttribute(int colSplit, int rowSplit) : this(colSplit, rowSplit, 0, 1)
    {
    }

    public FreezeAttribute(int colSplit, int rowSplit, int leftmostColumn, int topRow) =>
        FreezeSetting = new FreezeSetting(colSplit, rowSplit, leftmostColumn, topRow);

    internal FreezeSetting FreezeSetting { get; }

    /// <summary>
    ///     Horizontal position of split
    /// </summary>
    public int ColSplit => FreezeSetting.ColSplit;

    /// <summary>
    ///     Vertical position of split
    /// </summary>
    public int RowSplit => FreezeSetting.RowSplit;

    /// <summary>
    ///     Top row visible in bottom pane
    /// </summary>
    public int LeftMostColumn => FreezeSetting.LeftMostColumn;

    /// <summary>
    ///     Left column visible in right pane
    /// </summary>
    public int TopRow => FreezeSetting.TopRow;
}
