// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System.Text;

namespace WeihanLi.Npoi.Configurations;

public sealed class CsvOptions
{
    public string SeparatorString => new(SeparatorCharacter, 1);
    public string QuoteString => new(QuoteCharacter, 1);
    public char SeparatorCharacter { get; set; }
    public char QuoteCharacter { get; set; }
    public bool IncludeHeader { get; set; }
    public string PropertyNameForBasicType { get; set; }
    public Encoding Encoding { get; set; }

    public CsvOptions()
    {
        SeparatorCharacter = CsvHelper.CsvSeparatorCharacter;
        QuoteCharacter = CsvHelper.CsvQuoteCharacter;
        IncludeHeader = true;
        PropertyNameForBasicType = InternalConstants.DefaultPropertyNameForBasicType;
        Encoding = Encoding.UTF8;
    }

    public static readonly CsvOptions Default = new()
    {
        SeparatorCharacter = ',',
        QuoteCharacter = '"',
        PropertyNameForBasicType = InternalConstants.DefaultPropertyNameForBasicType
    };
}
