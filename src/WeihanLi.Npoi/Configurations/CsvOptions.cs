// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System.Text;

namespace WeihanLi.Npoi.Configurations;

/// <summary>
/// Represents the configurable behaviors of the CSV helper.
/// </summary>
public sealed class CsvOptions
{
    /// <summary>
    ///     Gets the separator character expressed as a string.
    /// </summary>
    public string SeparatorString => new(SeparatorCharacter, 1);

    /// <summary>
    ///     Gets the quote character expressed as a string.
    /// </summary>
    public string QuoteString => new(QuoteCharacter, 1);

    /// <summary>
    ///     Gets or sets the separator used between values.
    /// </summary>
    public char SeparatorCharacter { get; set; }

    /// <summary>
    ///     Gets or sets the character used to wrap textual values.
    /// </summary>
    public char QuoteCharacter { get; set; }

    /// <summary>
    ///     Gets or sets whether the header row should be emitted.
    /// </summary>
    public bool IncludeHeader { get; set; }

    /// <summary>
    ///     Gets or sets the synthetic property name to use for basic types.
    /// </summary>
    public string PropertyNameForBasicType { get; set; }

    /// <summary>
    ///     Gets or sets the encoding of the generated CSV.
    /// </summary>
    public Encoding Encoding { get; set; }

    /// <summary>
    ///     Initializes options with default separator, quote, and encoding.
    /// </summary>
    public CsvOptions()
    {
        SeparatorCharacter = CsvHelper.CsvSeparatorCharacter;
        QuoteCharacter = CsvHelper.CsvQuoteCharacter;
        IncludeHeader = true;
        PropertyNameForBasicType = InternalConstants.DefaultPropertyNameForBasicType;
        Encoding = Encoding.UTF8;
    }
    /// <summary>
    ///     Provides a shared instance representing sensible defaults.
    /// </summary>
    public static readonly CsvOptions Default = new()
    {
        SeparatorCharacter = ',',
        QuoteCharacter = '"',
        PropertyNameForBasicType = InternalConstants.DefaultPropertyNameForBasicType
    };
}
