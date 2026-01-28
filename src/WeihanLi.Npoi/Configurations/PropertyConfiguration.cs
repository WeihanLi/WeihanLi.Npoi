// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using NPOI.SS.UserModel;
using System.Reflection;
using WeihanLi.Extensions;

namespace WeihanLi.Npoi.Configurations;

internal class PropertyConfiguration : IPropertyConfiguration
{
    /// <summary>
    ///     ColumnIndex
    /// </summary>
    public int ColumnIndex { get; set; } = -1;

    /// <summary>
    ///     ColumnWidth
    /// </summary>
    public int ColumnWidth { get; set; }

    /// <summary>
    ///     Title
    /// </summary>
    public string ColumnTitle { get; set; } = string.Empty;

    /// <summary>
    ///     Formatter
    /// </summary>
    public string? ColumnFormatter { get; set; }

    /// <summary>
    ///     the property is ignored.
    /// </summary>
    public bool IsIgnored { get; set; }

    /// <summary>
    ///     PropertyName
    /// </summary>
    public string? PropertyName { get; set; }
}

internal sealed class PropertyConfiguration<TEntity, TProperty> : PropertyConfiguration,
    IPropertyConfiguration<TEntity, TProperty>
{
    private readonly PropertyInfo _propertyInfo;

    /// <summary>
    ///     Initializes a configuration wrapper for the specified property.
    /// </summary>
    /// <param name="propertyInfo">Property metadata.</param>
    public PropertyConfiguration(PropertyInfo propertyInfo)
    {
        _propertyInfo = propertyInfo;
        PropertyName = propertyInfo.Name;
        ColumnTitle = propertyInfo.Name;
    }

    /// <summary>
    ///     Sets the column index explicitly.
    /// </summary>
    public IPropertyConfiguration<TEntity, TProperty> HasColumnIndex(int index)
    {
        if (index >= 0)
        {
            ColumnIndex = index;
        }

        return this;
    }

    /// <summary>
    ///     Assigns the header text used for the column.
    /// </summary>
    public IPropertyConfiguration<TEntity, TProperty> HasColumnTitle(string title)
    {
        ColumnTitle = title ?? throw new ArgumentNullException(nameof(title));
        return this;
    }

    /// <summary>
    ///     Sets the column width (characters).
    /// </summary>
    public IPropertyConfiguration<TEntity, TProperty> HasColumnWidth(int width)
    {
        ColumnWidth = width;
        return this;
    }

    /// <summary>
    ///     Assigns the formatter string used when writing out the column.
    /// </summary>
    public IPropertyConfiguration<TEntity, TProperty> HasColumnFormatter(string? formatter)
    {
        ColumnFormatter = formatter;
        return this;
    }

    /// <summary>
    ///     Marks the property as ignored when exporting/importing.
    /// </summary>
    public IPropertyConfiguration<TEntity, TProperty> Ignored(bool ignored = true)
    {
        IsIgnored = ignored;
        return this;
    }

    /// <summary>
    ///     Registers a custom cell reader for imports.
    /// </summary>
    public IPropertyConfiguration<TEntity, TProperty> HasCellReader(Func<ICell, TProperty>? cellReader)
    {
        InternalCache.CellReaderFuncCache.AddOrUpdate(_propertyInfo, cellReader);
        return this;
    }

    /// <summary>
    ///     Registers a formatter used when exporting the property.
    /// </summary>
    public IPropertyConfiguration<TEntity, TProperty> HasOutputFormatter(
        Func<TEntity?, TProperty?, object?>? formatterFunc)
    {
        InternalCache.OutputFormatterFuncCache.AddOrUpdate(_propertyInfo, formatterFunc);
        return this;
    }

    /// <summary>
    ///     Registers a formatter used when importing cell values.
    /// </summary>
    public IPropertyConfiguration<TEntity, TProperty> HasInputFormatter(
        Func<TEntity?, TProperty?, TProperty?>? formatterFunc)
    {
        InternalCache.InputFormatterFuncCache.AddOrUpdate(_propertyInfo, formatterFunc);
        return this;
    }

    /// <summary>
    ///     Registers a formatter that manipulates the raw column text before parsing.
    /// </summary>
    public IPropertyConfiguration<TEntity, TProperty> HasColumnInputFormatter(
        Func<string?, TProperty?>? formatterFunc)
    {
        InternalCache.ColumnInputFormatterFuncCache.AddOrUpdate(_propertyInfo, formatterFunc);
        return this;
    }
}
