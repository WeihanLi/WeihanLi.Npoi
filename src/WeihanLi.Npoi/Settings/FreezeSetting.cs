// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

namespace WeihanLi.Npoi.Settings;

internal sealed class FreezeSetting
{
    /// <summary>
    ///     Initializes a freeze pane using default anchors.
    /// </summary>
    public FreezeSetting(int colSplit, int rowSplit) : this(colSplit, rowSplit, 0, 1)
    {
    }

    /// <summary>
    ///     Initializes a freeze pane with explicit anchors.
    /// </summary>
    /// <param name="colSplit">Horizontal split position.</param>
    /// <param name="rowSplit">Vertical split position.</param>
    /// <param name="leftmostColumn">Leftmost column displayed in the right pane.</param>
    /// <param name="topRow">Top row displayed in the bottom pane.</param>
    public FreezeSetting(int colSplit, int rowSplit, int leftmostColumn, int topRow)
    {
        ColSplit = colSplit;
        RowSplit = rowSplit;
        LeftMostColumn = leftmostColumn;
        TopRow = topRow;
    }

    /// <summary>
    ///     horizontal position of split
    /// </summary>
    public int ColSplit { get; }

    /// <summary>
    ///     Vertical position of split
    /// </summary>
    public int RowSplit { get; }

    /// <summary>
    ///     Top row visible in bottom pane
    /// </summary>
    public int LeftMostColumn { get; }

    /// <summary>
    ///     Left column visible in right pane
    /// </summary>
    public int TopRow { get; }
}
