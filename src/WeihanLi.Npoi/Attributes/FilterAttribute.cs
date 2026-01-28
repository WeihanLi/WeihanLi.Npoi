// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Attributes;

/// <summary>
/// Specifies the auto-filter range that should be applied to a sheet.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
public sealed class FilterAttribute : Attribute
{
    /// <summary>
    ///     Initializes the attribute for a single column.
    /// </summary>
    public FilterAttribute(int firstColumn) : this(firstColumn, null)
    {
    }

    /// <summary>
    ///     Initializes the attribute for an explicit column range.
    /// </summary>
    /// <param name="firstColumn">First column index.</param>
    /// <param name="lastColumn">Last column index.</param>
    public FilterAttribute(int firstColumn, int? lastColumn) =>
        FilterSetting = new FilterSetting(firstColumn, lastColumn);

    internal FilterSetting FilterSetting { get; }

    /// <summary>
    ///     Gets or sets the first column index.
    /// </summary>
    public int FirstColumn => FilterSetting.FirstColumn;

    /// <summary>
    ///     Gets or sets the last column index.
    /// </summary>
    public int? LastColumn => FilterSetting.LastColumn;
}
