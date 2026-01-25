// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

namespace WeihanLi.Npoi.Settings;

internal sealed class FilterSetting
{
    /// <summary>
    ///     Initializes a filter specification.
    /// </summary>
    /// <param name="firstColumn">First column index.</param>
    /// <param name="lastColumn">Optional last column index.</param>
    public FilterSetting(int firstColumn, int? lastColumn)
    {
        FirstColumn = firstColumn;
        LastColumn = lastColumn;
    }

    /// <summary>
    ///     Gets or sets the first column index.
    /// </summary>
    public int FirstColumn { get; }

    /// <summary>
    ///     Gets or sets the last column index.
    /// </summary>
    public int? LastColumn { get; }
}
