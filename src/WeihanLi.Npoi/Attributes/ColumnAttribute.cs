using System;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi.Attributes;

[AttributeUsage(AttributeTargets.Property)]
public sealed class ColumnAttribute : Attribute
{
    public ColumnAttribute() => PropertyConfiguration = new PropertyConfiguration();

    public ColumnAttribute(int index) => PropertyConfiguration = new PropertyConfiguration { ColumnIndex = index };

    public ColumnAttribute(string title) => PropertyConfiguration = new PropertyConfiguration
    {
        ColumnTitle = title ?? throw new ArgumentNullException(nameof(title))
    };

    internal PropertyConfiguration PropertyConfiguration { get; }

    /// <summary>
    ///     ColumnIndex
    /// </summary>
    public int Index
    {
        get => PropertyConfiguration.ColumnIndex;
        set
        {
            if (value >= 0)
            {
                PropertyConfiguration.ColumnIndex = value;
            }
        }
    }

    /// <summary>
    ///     ColumnTitle
    /// </summary>
    public string Title
    {
        get => PropertyConfiguration.ColumnTitle;
        set => PropertyConfiguration.ColumnTitle = value;
    }

    /// <summary>
    ///     Formatter
    /// </summary>
    public string? Formatter
    {
        get => PropertyConfiguration.ColumnFormatter;
        set => PropertyConfiguration.ColumnFormatter = value;
    }

    /// <summary>
    ///     IsIgnored
    /// </summary>
    public bool IsIgnored
    {
        get => PropertyConfiguration.IsIgnored;
        set => PropertyConfiguration.IsIgnored = value;
    }

    /// <summary>
    ///     ColumnWidth
    ///     Characters Count
    /// </summary>
    public int Width
    {
        get => PropertyConfiguration.ColumnWidth;
        set => PropertyConfiguration.ColumnWidth = value;
    }
}
