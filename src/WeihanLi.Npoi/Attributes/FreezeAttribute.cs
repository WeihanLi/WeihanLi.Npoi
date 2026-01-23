// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Attributes;

/// <summary>
/// Declares a freeze pane for a mapped sheet.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
public sealed class FreezeAttribute : Attribute
{
    /// <summary>
    ///     Initializes a freeze pane using default anchors.
    /// </summary>
    public FreezeAttribute(int colSplit, int rowSplit) : this(colSplit, rowSplit, 0, 1)
    {
    }

    /// <summary>
    ///     Initializes a freeze pane with explicit anchors.
    /// </summary>
    /// <param name="colSplit">Horizontal split position.</param>
    /// <param name="rowSplit">Vertical split position.</param>
    /// <param name="leftmostColumn">Left column visible in right pane.</param>
    /// <param name="topRow">Top row visible in bottom pane.</param>
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
