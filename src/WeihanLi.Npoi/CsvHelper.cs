// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using WeihanLi.Common;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi;

public sealed class CsvOptions
{
    public char SeparatorCharacter { get; set; }
    public char QuoteCharacter { get; set; }
    public bool IncludeHeader { get; set; }
    public string PropertyNameForBasicType { get; set; }

    public CsvOptions()
    {
        SeparatorCharacter = CsvHelper.CsvSeparatorCharacter;
        QuoteCharacter = CsvHelper.CsvQuoteCharacter;
        IncludeHeader = true;
        PropertyNameForBasicType = InternalConstants.DefaultPropertyNameForBasicType;
    }
}

/// <summary>
///     CsvHelper
/// </summary>
public static class CsvHelper
{
    /// <summary>
    ///     CsvSeparatorCharacter, ',' by default
    /// </summary>
    public static char CsvSeparatorCharacter = ',';

    /// <summary>
    ///     CsvQuoteCharacter, '"' by default
    /// </summary>
    public static char CsvQuoteCharacter = '"';

    /// <summary>
    ///     save to csv file
    /// </summary>
    public static bool ToCsvFile(this DataTable dt, string filePath) => ToCsvFile(dt, filePath, true);

    /// <summary>
    ///     save to csv file
    /// </summary>
    public static bool ToCsvFile(this DataTable dataTable, string filePath, bool includeHeader)
    {
        return ToCsvFile(dataTable, filePath, new CsvOptions() { IncludeHeader = includeHeader });
    }
    
    /// <summary>
    ///     save to csv file
    /// </summary>
    public static bool ToCsvFile(this DataTable dataTable, string filePath, CsvOptions csvOptions)
    {
        if (dataTable is null)
        {
            throw new ArgumentNullException(nameof(dataTable));
        }

        var dir = Path.GetDirectoryName(filePath);
        if (dir is not null)
        {
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        var csvText = GetCsvText(dataTable, csvOptions);
        if (csvText.IsNullOrWhiteSpace())
        {
            return false;
        }

        File.WriteAllText(filePath, csvText, Encoding.UTF8);
        return true;
    }

    /// <summary>
    ///     to csv bytes
    /// </summary>
    public static byte[] ToCsvBytes(this DataTable dt) => ToCsvBytes(dt, true);

    /// <summary>
    ///     to csv bytes
    /// </summary>
    public static byte[] ToCsvBytes(this DataTable dataTable, bool includeHeader) =>
        GetCsvText(dataTable, includeHeader).GetBytes();
    
    /// <summary>
    ///     to csv bytes
    /// </summary>
    public static byte[] ToCsvBytes(this DataTable dataTable, CsvOptions csvOptions) =>
        GetCsvText(dataTable, csvOptions).GetBytes();

    /// <summary>
    ///     convert csv file data to dataTable
    /// </summary>
    /// <param name="csvBytes">csv bytes</param>
    public static DataTable ToDataTable(byte[] csvBytes)
        => ToDataTable(csvBytes, new CsvOptions());
    
    public static DataTable ToDataTable(byte[] csvBytes, CsvOptions csvOptions)
    {
        if (csvBytes is null)
        {
            throw new ArgumentNullException(nameof(csvBytes));
        }

        using var ms = new MemoryStream(csvBytes);
        return ToDataTable(ms, csvOptions);
    }

    /// <summary>
    ///     convert csv stream data to dataTable
    /// </summary>
    /// <param name="stream">stream</param>
    public static DataTable ToDataTable(Stream stream) => ToDataTable(stream, new CsvOptions());
    
    public static DataTable ToDataTable(Stream stream, CsvOptions csvOptions)
    {
        if (stream is null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        var dt = new DataTable();

        if (stream.CanRead)
        {
            using var sr = new StreamReader(stream, Encoding.UTF8);
            string strLine;
            var isFirst = true;
            while ((strLine = sr.ReadLine()!).IsNotNullOrEmpty())
            {
                var rowData = ParseLine(strLine, csvOptions);
                var dtColumns = rowData.Count;
                if (isFirst)
                {
                    for (var i = 0; i < dtColumns; i++)
                    {
                        var columnName = rowData[i];
                        if (dt.Columns.Contains(columnName))
                        {
                            columnName = InternalHelper.GetEncodedColumnName(columnName);
                        }

                        dt.Columns.Add(columnName);
                    }

                    isFirst = false;
                }
                else
                {
                    var dataRow = dt.NewRow();
                    for (var j = 0; j < dt.Columns.Count; j++)
                    {
                        dataRow[j] = rowData[j];
                    }

                    dt.Rows.Add(dataRow);
                }
            }
        }

        return dt;
    }

    /// <summary>
    ///     convert csv file data to dataTable
    /// </summary>
    /// <param name="filePath">csv file path</param>
    public static DataTable ToDataTable(string filePath)
    {
        if (filePath is null)
        {
            throw new ArgumentNullException(nameof(filePath));
        }

        if (!File.Exists(filePath))
        {
            throw new ArgumentException(Resource.FileNotFound, nameof(filePath));
        }

        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ToDataTable(fs);
    }

    /// <summary>
    ///     convert csv file data to entity list
    /// </summary>
    /// <param name="filePath">csv file path</param>
    public static List<TEntity?> ToEntityList<TEntity>(string filePath)
        => ToEntityList<TEntity>(filePath, new CsvOptions());
    
    public static List<TEntity?> ToEntityList<TEntity>(string filePath, CsvOptions csvOptions)
    {
        Guard.NotNull(filePath);
        if (!File.Exists(filePath))
        {
            throw new ArgumentException(Resource.FileNotFound, nameof(filePath));
        }
        return ToEntityList<TEntity?>(File.ReadAllBytes(filePath), csvOptions);
    }

    /// <summary>
    ///     convert csv file data to entity list
    /// </summary>
    /// <param name="csvBytes">csv bytes</param>
    public static List<TEntity?> ToEntityList<TEntity>(byte[] csvBytes)
        => ToEntityList<TEntity>(csvBytes, new CsvOptions());
    
    public static List<TEntity?> ToEntityList<TEntity>(byte[] csvBytes, CsvOptions csvOptions)
    {
        Guard.NotNull(csvBytes);
        using var ms = new MemoryStream(csvBytes);
        return ToEntityList<TEntity>(ms, csvOptions);
    }

    /// <summary>
    ///     convert csv file data to entity list
    /// </summary>
    /// <param name="csvStream">csv Stream</param>
    public static List<TEntity?> ToEntityList<TEntity>(Stream csvStream)
        => ToEntityList<TEntity>(csvStream, new CsvOptions());
    
    public static List<TEntity?> ToEntityList<TEntity>(Stream csvStream, CsvOptions csvOptions)
    {
        Guard.NotNull(csvStream);
        csvStream.Seek(0, SeekOrigin.Begin);
        
        var entities = new List<TEntity?>();
        if (typeof(TEntity).IsBasicType())
        {
            using var sr = new StreamReader(csvStream, Encoding.UTF8);
            string strLine;
            var isFirstLine = true;
            while ((strLine = sr.ReadLine()!).IsNotNullOrEmpty())
            {
                if (isFirstLine)
                {
                    isFirstLine = false;
                    continue;
                }

                //
                entities.Add(strLine.Trim().To<TEntity>());
            }
        }
        else
        {
            var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();
            var propertyColumnDictionary = InternalHelper.GetPropertyColumnDictionary<TEntity>();
            var propertyColumnDic = propertyColumnDictionary.ToDictionary(_ => _.Key,
                _ => new PropertyConfiguration
                {
                    ColumnIndex = -1,
                    ColumnFormatter = _.Value.ColumnFormatter,
                    ColumnTitle = _.Value.ColumnTitle,
                    ColumnWidth = _.Value.ColumnWidth,
                    IsIgnored = _.Value.IsIgnored
                });
            using var sr = new StreamReader(csvStream, Encoding.UTF8);
            string strLine;
            var isFirstLine = true;
            while ((strLine = sr.ReadLine()!).IsNotNullOrEmpty())
            {
                var entityType = typeof(TEntity);
                var cols = ParseLine(strLine, csvOptions);
                if (isFirstLine)
                {
                    for (var index = 0; index < cols.Count; index++)
                    {
                        var setting = propertyColumnDic.GetPropertySetting(cols[index]);
                        if (setting != null)
                        {
                            setting.ColumnIndex = index;
                        }
                    }

                    if (propertyColumnDic.Values.All(_ => _.ColumnIndex < 0))
                    {
                        propertyColumnDic = propertyColumnDictionary;
                    }

                    isFirstLine = false;
                }
                else
                {
                    var entity = NewFuncHelper<TEntity>.Instance();
                    if (entityType.IsValueType)
                    {
                        var obj = (object)entity!; // boxing for value types

                        foreach (var key in propertyColumnDic.Keys)
                        {
                            var colIndex = propertyColumnDic[key].ColumnIndex;
                            if (colIndex >= 0 && colIndex < cols.Count && key.CanWrite)
                            {
                                var columnValue = key.PropertyType.GetDefaultValue();
                                var valueApplied = false;
                                if (InternalCache.ColumnInputFormatterFuncCache.TryGetValue(key,
                                    out var formatterFunc) && formatterFunc?.Method != null)
                                {
                                    var cellValue = cols[colIndex];
                                    try
                                    {
                                        // apply custom formatterFunc
                                        columnValue = formatterFunc.DynamicInvoke(cellValue);
                                        valueApplied = true;
                                    }
                                    catch (Exception e)
                                    {
                                        Debug.WriteLine(e);
                                        InvokeHelper.OnInvokeException?.Invoke(e);
                                    }
                                }

                                if (valueApplied == false)
                                {
                                    columnValue = cols[colIndex].ToOrDefault(key.PropertyType);
                                }

                                key.GetValueSetter()?.Invoke(entity!, columnValue);
                            }
                        }

                        entity = (TEntity)obj; // unboxing
                    }
                    else
                    {
                        foreach (var key in propertyColumnDic.Keys)
                        {
                            var colIndex = propertyColumnDic[key].ColumnIndex;
                            if (colIndex >= 0 && colIndex < cols.Count && key.CanWrite)
                            {
                                var columnValue = key.PropertyType.GetDefaultValue();

                                var valueApplied = false;
                                if (InternalCache.ColumnInputFormatterFuncCache.TryGetValue(key,
                                    out var formatterFunc) && formatterFunc?.Method != null)
                                {
                                    var cellValue = cols[colIndex];
                                    try
                                    {
                                        // apply custom formatterFunc
                                        columnValue = formatterFunc.DynamicInvoke(cellValue);
                                        valueApplied = true;
                                    }
                                    catch (Exception e)
                                    {
                                        Debug.WriteLine(e);
                                        InvokeHelper.OnInvokeException?.Invoke(e);
                                    }
                                }

                                if (valueApplied == false)
                                {
                                    columnValue = cols[colIndex].ToOrDefault(key.PropertyType);
                                }

                                key.GetValueSetter()?.Invoke(entity!, columnValue);
                            }
                        }
                    }

                    if (null != entity)
                    {
                        foreach (var propertyInfo in propertyColumnDic.Keys)
                        {
                            if (propertyInfo.CanWrite)
                            {
                                var propertyValue = propertyInfo.GetValueGetter()?.Invoke(entity);
                                if (InternalCache.InputFormatterFuncCache.TryGetValue(propertyInfo,
                                    out var formatterFunc) && formatterFunc?.Method != null)
                                {
                                    try
                                    {
                                        // apply custom formatterFunc
                                        var formattedValue = formatterFunc.DynamicInvoke(entity, propertyValue);
                                        propertyInfo.GetValueSetter()?.Invoke(entity, formattedValue);
                                    }
                                    catch (Exception e)
                                    {
                                        Debug.WriteLine(e);
                                        InvokeHelper.OnInvokeException?.Invoke(e);
                                    }
                                }
                            }
                        }
                    }

                    if (configuration.DataValidationFunc != null && !configuration.DataValidationFunc(entity))
                    {
                        // data invalid
                        continue;
                    }

                    entities.Add(entity);
                }
            }
        }

        return entities;
    }

    public static IReadOnlyList<string> ParseLine(string line) => ParseLine(line, new CsvOptions());
    
    public static IReadOnlyList<string> ParseLine(string line, CsvOptions csvOptions)
    {
        if (string.IsNullOrEmpty(line))
        {
            return Array.Empty<string>();
        }

        var columnBuilder = new StringBuilder();
        var fields = new List<string>();

        var inColumn = false;
        var inQuotes = false;

        // Iterate through every character in the line
        for (var i = 0; i < line.Length; i++)
        {
            var character = line[i];

            // If we are not currently inside a column
            if (!inColumn)
            {
                // If the current character is a double quote then the column value is contained within
                // double quotes, otherwise append the next character
                inColumn = true;
                if (character == csvOptions.QuoteCharacter)
                {
                    inQuotes = true;
                    continue;
                }
            }

            // If we are in between double quotes
            if (inQuotes)
            {
                if (i + 1 == line.Length)
                {
                    break;
                }

                if (character == csvOptions.QuoteCharacter && line[i + 1] == csvOptions.SeparatorCharacter) // quotes end
                {
                    inQuotes = false;
                    inColumn = false;
                    i++; //skip next
                }
                else if (character == csvOptions.QuoteCharacter && line[i + 1] == csvOptions.QuoteCharacter) // quotes
                {
                    i++; //skip next
                }
                else if (character == csvOptions.QuoteCharacter)
                {
                    throw new ArgumentException($"unable to escape {line}");
                }
            }
            else if (character == csvOptions.SeparatorCharacter)
            {
                inColumn = false;
            }

            // If we are no longer in the column clear the builder and add the columns to the list
            if (!inColumn)
            {
                fields.Add(columnBuilder.ToString());
                columnBuilder.Clear();
            }
            else // append the current column
            {
                columnBuilder.Append(character);
            }
        }

        fields.Add(columnBuilder.ToString());

        return fields;
    }

    /// <summary>
    ///     save to csv file
    /// </summary>
    public static bool ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath) =>
        ToCsvFile(entities, filePath, true);

    /// <summary>
    ///     save to csv file
    /// </summary>
    public static bool ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath, bool includeHeader)
    {
        return ToCsvFile(Guard.NotNull(entities), filePath, new CsvOptions() { IncludeHeader = includeHeader });
    }
    
    public static bool ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath, CsvOptions csvOptions)
    {
        if (entities is null)
        {
            throw new ArgumentNullException(nameof(entities));
        }

        var dir = Path.GetDirectoryName(filePath);
        if (dir is not null)
        {
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }
        
        var csvTextData = GetCsvText(entities, csvOptions);
        if (csvTextData.IsNullOrEmpty())
        {
            return false;
        }

        File.WriteAllText(filePath, csvTextData, Encoding.UTF8);
        return true;
    }

    /// <summary>
    ///     to csv bytes
    /// </summary>
    public static byte[] ToCsvBytes<TEntity>(this IEnumerable<TEntity> entities) => ToCsvBytes(entities, true);

    /// <summary>
    ///     to csv bytes
    /// </summary>
    public static byte[] ToCsvBytes<TEntity>(this IEnumerable<TEntity> entities, bool includeHeader) =>
        GetCsvText(entities, includeHeader).GetBytes();
    
    /// <summary>
    ///     to csv bytes
    /// </summary>
    public static byte[] ToCsvBytes<TEntity>(this IEnumerable<TEntity> entities, CsvOptions csvOptions) =>
        GetCsvText(entities, csvOptions).GetBytes();

    /// <summary>
    ///     Get csv text
    /// </summary>
    public static string GetCsvText<TEntity>(this IEnumerable<TEntity> entities, bool includeHeader = true)
    {
        return GetCsvText(Guard.NotNull(entities), new CsvOptions()
        {
            IncludeHeader = includeHeader
        });
    }

    /// <summary>
    ///     Get csv text
    /// </summary>
    public static string GetCsvText<TEntity>(this IEnumerable<TEntity> entities, CsvOptions csvOptions)
    {
        if (entities is null)
        {
            throw new ArgumentNullException(nameof(entities));
        }
        Guard.NotNull(csvOptions);

        var data = new StringBuilder();
        var isBasicType = typeof(TEntity).IsBasicType();
        if (isBasicType)
        {
            if (csvOptions.IncludeHeader)
            {
                data.AppendLine(csvOptions.PropertyNameForBasicType);
            }

            foreach (var entity in entities)
            {
                data.AppendLine(Convert.ToString(entity));
            }
        }
        else
        {
            var dic = InternalHelper.GetPropertyColumnDictionary<TEntity>();
            var props = InternalHelper.GetPropertiesForCsvHelper<TEntity>();
            if (csvOptions.IncludeHeader)
            {
                for (var i = 0; i < props.Count; i++)
                {
                    if (i > 0)
                    {
                        data.Append(csvOptions.SeparatorCharacter);
                    }

                    data.Append(dic[props[i]].ColumnTitle);
                }

                data.AppendLine();
            }

            foreach (var entity in entities)
            {
                for (var i = 0; i < props.Count; i++)
                {
                    var propertyValue = props[i].GetValueGetter<TEntity>()?.Invoke(entity);
                    if (InternalCache.OutputFormatterFuncCache.TryGetValue(props[i], out var formatterFunc) &&
                        formatterFunc?.Method != null)
                    {
                        try
                        {
                            // apply custom formatterFunc
                            propertyValue = formatterFunc.DynamicInvoke(entity, propertyValue);
                        }
                        catch (Exception e)
                        {
                            Debug.WriteLine(e);
                            InvokeHelper.OnInvokeException?.Invoke(e);
                        }
                    }

                    if (i > 0)
                    {
                        data.Append(csvOptions.SeparatorCharacter);
                    }

                    // https://stackoverflow.com/questions/4617935/is-there-a-way-to-include-commas-in-csv-columns-without-breaking-the-formatting
                    var val = propertyValue?.ToString()?.Replace("\"", "\"\"");
                    if (val is { Length: > 0 })
                    {
                        data.Append(val.IndexOf(csvOptions.SeparatorCharacter) > -1 ? $"\"{val}\"" : val);
                    }
                }

                data.AppendLine();
            }
        }

        return data.ToString();
    }

    /// <summary>
    ///     Get csv text
    /// </summary>
    public static string GetCsvText(this DataTable? dataTable, bool includeHeader = true)
    {
        return GetCsvText(dataTable, new CsvOptions()
        {
            IncludeHeader = includeHeader
        });
    }
    
    public static string GetCsvText(this DataTable? dataTable, CsvOptions csvOptions)
    {
        Guard.NotNull(csvOptions);
        if (dataTable is null || dataTable.Rows.Count == 0 || dataTable.Columns.Count == 0)
        {
            return string.Empty;
        }

        var data = new StringBuilder();

        if (csvOptions.IncludeHeader)
        {
            for (var i = 0; i < dataTable.Columns.Count; i++)
            {
                if (i > 0)
                {
                    data.Append(csvOptions.SeparatorCharacter);
                }

                var columnName = InternalHelper.GetDecodeColumnName(dataTable.Columns[i].ColumnName);
                data.Append(columnName);
            }

            data.AppendLine();
        }

        for (var i = 0; i < dataTable.Rows.Count; i++)
        {
            for (var j = 0; j < dataTable.Columns.Count; j++)
            {
                if (j > 0)
                {
                    data.Append(csvOptions.SeparatorCharacter);
                }

                // https://stackoverflow.com/questions/4617935/is-there-a-way-to-include-commas-in-csv-columns-without-breaking-the-formatting
                var val = dataTable.Rows[i][j]?.ToString()?.Replace("\"", "\"\"");
                if (val is { Length: > 0 })
                {
                    data.Append(val.IndexOf(csvOptions.SeparatorCharacter) > -1 ? $"\"{val}\"" : val);
                }
            }

            data.AppendLine();
        }

        return data.ToString();
    }
}
