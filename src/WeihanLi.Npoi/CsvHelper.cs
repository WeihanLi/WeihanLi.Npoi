// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System.Data;
using System.Diagnostics;
using System.Text;
using WeihanLi.Common;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi;

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
    public static bool ToCsvFile(this DataTable dt, string filePath) => ToCsvFile(dt, filePath, CsvOptions.Default);

    /// <summary>
    ///     save to csv file
    /// </summary>
    public static bool ToCsvFile(this DataTable dataTable, string filePath, bool includeHeader)
    {
        return ToCsvFile(dataTable, filePath, includeHeader ? CsvOptions.Default : new CsvOptions()
        {
            IncludeHeader = false
        });
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

        var csvText = GetCsvText(dataTable, csvOptions);
        if (csvText.IsNullOrEmpty())
        {
            return false;
        }
        InternalHelper.EnsureFileIsNotReadOnly(filePath);
        var dir = Path.GetDirectoryName(filePath);
        if (dir.IsNotNullOrEmpty())
        {
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        File.WriteAllText(filePath, csvText, csvOptions.Encoding);
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
        => ToDataTable(csvBytes, CsvOptions.Default);

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
    public static DataTable ToDataTable(Stream stream) => ToDataTable(stream, CsvOptions.Default);

    /// <summary>
    ///     convert csv stream data to dataTable
    /// </summary>
    /// <param name="stream">stream</param>
    /// <param name="csvOptions">csvOptions</param>
    public static DataTable ToDataTable(Stream stream, CsvOptions csvOptions)
    {
        Guard.NotNull(stream);
        Guard.NotNull(csvOptions);

        var dt = new DataTable();

        if (stream.CanRead)
        {
            using var sr = new StreamReader(stream, csvOptions.Encoding);
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
        => ToEntityList<TEntity>(filePath, CsvOptions.Default);

    /// <summary>
    ///     convert csv file data to entity list
    /// </summary>
    /// <param name="filePath">csv file path</param>
    /// <param name="csvOptions">csvOptions</param>
    public static List<TEntity?> ToEntityList<TEntity>(string filePath, CsvOptions csvOptions)
    {
        Guard.NotNull(filePath);
        if (!File.Exists(filePath))
        {
            throw new ArgumentException(Resource.FileNotFound, nameof(filePath));
        }

        return GetEntityList<TEntity>(File.ReadAllLines(filePath), csvOptions);
    }

    /// <summary>
    ///     convert csv file data to entities
    /// </summary>
    /// <param name="filePath">csv file path</param>
    /// <param name="csvOptions">csvOptions</param>
    public static IEnumerable<TEntity?> ToEntities<TEntity>(string filePath, CsvOptions? csvOptions = null)
    {
        Guard.NotNull(filePath);
        if (!File.Exists(filePath))
        {
            throw new ArgumentException(Resource.FileNotFound, nameof(filePath));
        }
        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ToEntities<TEntity>(fs, csvOptions);
    }

    /// <summary>
    ///     convert csv byte data to entity list
    /// </summary>
    /// <param name="csvBytes">csv bytes</param>
    public static List<TEntity?> ToEntityList<TEntity>(byte[] csvBytes)
        => ToEntityList<TEntity>(csvBytes, CsvOptions.Default);

    /// <summary>
    ///     convert csv byte data to entity list
    /// </summary>
    /// <param name="csvBytes">csv bytes</param>
    /// <param name="csvOptions">csvOptions</param>
    public static List<TEntity?> ToEntityList<TEntity>(byte[] csvBytes, CsvOptions csvOptions)
    {
        Guard.NotNull(csvBytes);
        using var ms = new MemoryStream(csvBytes);
        return ToEntityList<TEntity>(ms, csvOptions);
    }

    public static IEnumerable<TEntity?> ToEntities<TEntity>(byte[] csvBytes, CsvOptions? csvOptions = null)
    {
        Guard.NotNull(csvBytes);
        using var ms = new MemoryStream(csvBytes);
        return ToEntities<TEntity>(ms, csvOptions);
    }

    /// <summary>
    ///     convert csv stream data to entity list
    /// </summary>
    /// <param name="csvStream">csv Stream</param>
    public static List<TEntity?> ToEntityList<TEntity>(Stream csvStream)
        => ToEntityList<TEntity>(csvStream, CsvOptions.Default);

    public static List<TEntity?> ToEntityList<TEntity>(Stream csvStream, CsvOptions csvOptions)
    {
        Guard.NotNull(csvStream);
        csvStream.Seek(0, SeekOrigin.Begin);

        using var reader = new StreamReader(csvStream, csvOptions.Encoding);
        var lines = new List<string>();

        while (true)
        {
            var strLine = reader.ReadLine();
            if (strLine.IsNullOrEmpty())
                break;

            lines.Add(strLine);
        }

        return GetEntityList<TEntity>(lines, csvOptions);
    }

    public static IEnumerable<TEntity?> ToEntities<TEntity>(Stream csvStream, CsvOptions? csvOptions = null)
    {
        Guard.NotNull(csvStream);

        var lines = GetLines();
        return GetEntityList<TEntity>(lines, csvOptions);

        IEnumerable<string> GetLines()
        {
            csvStream.Seek(0, SeekOrigin.Begin);
            using var reader = new StreamReader(csvStream, csvOptions?.Encoding ?? Encoding.UTF8);
            while (true)
            {
                var strLine = reader.ReadLine();
                if (strLine.IsNullOrEmpty())
                    yield break;

                yield return strLine;
            }
        }
    }

    public static List<TEntity?> GetEntityList<TEntity>(string csvText, CsvOptions? csvOptions = null)
    {
        Guard.NotNull(csvText);
        var lines = GetLines();
        return GetEntityList<TEntity>(lines, csvOptions);

        IEnumerable<string> GetLines()
        {
            using var reader = new StringReader(csvText);
            while (true)
            {
                var strLine = reader.ReadLine();
                if (strLine.IsNullOrEmpty())
                    yield break;

                yield return strLine;
            }
        }
    }

    public static List<TEntity?> GetEntityList<TEntity>(IEnumerable<string> csvLines, CsvOptions? csvOptions = null)
        => GetEntities<TEntity>(csvLines, csvOptions).ToList();

    public static IEnumerable<TEntity?> GetEntities<TEntity>(IEnumerable<string> csvLines, CsvOptions? csvOptions = null)
    {
        if (csvLines is null)
        {
            throw new ArgumentNullException(nameof(csvLines));
        }
        csvOptions ??= CsvOptions.Default;
        var entityType = typeof(TEntity);
        if (entityType.IsBasicType())
        {
            var lines = csvOptions.IncludeHeader ? csvLines.Skip(1) : csvLines;
            foreach (var strLine in lines)
            {
                yield return strLine.To<TEntity>();
            }
        }
        else
        {
            var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();
            var propertyColumnDictionary = InternalHelper.GetPropertyColumnDictionary<TEntity>();
            var propertyColumnDic = csvOptions.IncludeHeader ? propertyColumnDictionary.ToDictionary(_ => _.Key,
                _ => new PropertyConfiguration
                {
                    ColumnIndex = -1,
                    ColumnFormatter = _.Value.ColumnFormatter,
                    ColumnTitle = _.Value.ColumnTitle,
                    ColumnWidth = _.Value.ColumnWidth,
                    IsIgnored = _.Value.IsIgnored
                }) : propertyColumnDictionary;
            var isFirstLine = csvOptions.IncludeHeader;
            foreach (var strLine in csvLines)
            {
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

                    if (propertyColumnDic.Values.Any(_ => _.ColumnIndex < 0))
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
                    if (configuration.DataFilter?.Invoke(entity) == false)
                    {
                        continue;
                    }
                    yield return entity;
                }
            }
        }
    }

    public static IReadOnlyList<string> ParseLine(string line) => ParseLine(line, CsvOptions.Default);

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
        ToCsvFile(entities, filePath, CsvOptions.Default);

    /// <summary>
    ///     save to csv file
    /// </summary>
    public static bool ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath, bool includeHeader)
    {
        return ToCsvFile(Guard.NotNull(entities), filePath, includeHeader ? CsvOptions.Default : new CsvOptions()
        {
            IncludeHeader = false
        });
    }

    public static bool ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath, CsvOptions csvOptions)
    {
        if (entities is null)
        {
            throw new ArgumentNullException(nameof(entities));
        }
        Guard.NotNull(csvOptions);

        var csvTextData = GetCsvText(entities, csvOptions);
        if (csvTextData.IsNullOrEmpty())
        {
            return false;
        }

        var dir = Path.GetDirectoryName(filePath);
        if (dir.IsNotNullOrEmpty())
        {
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        File.WriteAllText(filePath, csvTextData, csvOptions.Encoding);
        return true;
    }

#if NET6_0
    public static async Task<bool> ToCsvFileAsync<TEntity>(this IEnumerable<TEntity> entities, string filePath, CsvOptions? csvOptions = null)
    {
        if (entities is null)
        {
            throw new ArgumentNullException(nameof(entities));
        }

        csvOptions ??= CsvOptions.Default;
        var csvTextData = GetCsvText(entities, csvOptions);
        if (csvTextData.IsNullOrEmpty())
        {
            return false;
        }

        InternalHelper.EnsureFileIsNotReadOnly(filePath);
        var dir = Path.GetDirectoryName(filePath);
        if (dir.IsNotNullOrEmpty())
        {
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        await File.WriteAllTextAsync(filePath, csvTextData, csvOptions.Encoding);
        return true;
    }
#endif

    /// <summary>
    ///     to csv bytes
    /// </summary>
    public static byte[] ToCsvBytes<TEntity>(this IEnumerable<TEntity> entities) => ToCsvBytes(entities, CsvOptions.Default);

    /// <summary>
    ///     to csv bytes
    /// </summary>
    public static byte[] ToCsvBytes<TEntity>(this IEnumerable<TEntity> entities, bool includeHeader) =>
        GetCsvText(entities, includeHeader).GetBytes();

    /// <summary>
    ///     to csv bytes
    /// </summary>
    public static byte[] ToCsvBytes<TEntity>(this IEnumerable<TEntity> entities, CsvOptions csvOptions) =>
        GetCsvText(entities, csvOptions).GetBytes(csvOptions.Encoding);

    /// <summary>
    ///     Get csv text
    /// </summary>
    public static string GetCsvText<TEntity>(this IEnumerable<TEntity> entities, bool includeHeader = true)
    {
        return GetCsvText(Guard.NotNull(entities), includeHeader ? CsvOptions.Default : new CsvOptions()
        {
            IncludeHeader = false
        });
    }

    /// <summary>
    ///     Get csv text
    /// </summary>
    public static string GetCsvText<TEntity>(this IEnumerable<TEntity> entities, CsvOptions csvOptions) =>
        GetCsvLines(entities, csvOptions).StringJoin(Environment.NewLine);

    /// <summary>
    ///     Get csv lines
    /// </summary>
    /// <param name="entities">entities</param>
    /// <param name="csvOptions">csvOptions</param>
    /// <typeparam name="TEntity">entity type</typeparam>
    /// <returns>csv lines</returns>
    public static IEnumerable<string> GetCsvLines<TEntity>(this IEnumerable<TEntity> entities, CsvOptions? csvOptions = null)
    {
        if (entities is null)
        {
            throw new ArgumentNullException(nameof(entities));
        }
        csvOptions ??= CsvOptions.Default;

        var isBasicType = typeof(TEntity).IsBasicType();
        if (isBasicType)
        {
            if (csvOptions.IncludeHeader)
            {
                yield return csvOptions.PropertyNameForBasicType;
            }
            foreach (var entity in entities)
            {
                yield return Convert.ToString(entity) ?? string.Empty;
            }
        }
        else
        {
            var dic = InternalHelper.GetPropertyColumnDictionary<TEntity>();
            var props = InternalHelper.GetPropertiesForCsvHelper<TEntity>();
            if (csvOptions.IncludeHeader)
            {
                yield return Enumerable.Range(0, props.Count)
                    .Select(i => dic[props[i]].ColumnTitle)
                    .StringJoin(csvOptions.SeparatorString);
            }

            foreach (var entity in entities)
            {
                var line = GetCsvLine().StringJoin(csvOptions.SeparatorString);
                yield return line;

                IEnumerable<string> GetCsvLine()
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

                        // https://stackoverflow.com/questions/4617935/is-there-a-way-to-include-commas-in-csv-columns-without-breaking-the-formatting
                        var val = propertyValue?.ToString()?.Replace(
                            csvOptions.QuoteString,
                            $"{csvOptions.QuoteString}{csvOptions.QuoteString}"
                        );
                        if (val is { Length: > 0 })
                        {
                            yield return val.IndexOf(csvOptions.SeparatorCharacter) > -1 ? $"{csvOptions.QuoteCharacter}{val}{csvOptions.QuoteCharacter}" : val;
                        }
                        else
                        {
                            yield return string.Empty;
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    ///     Get csv text
    /// </summary>
    public static string GetCsvText(this DataTable? dataTable, bool includeHeader = true)
    {
        return GetCsvText(dataTable, includeHeader ? CsvOptions.Default : new CsvOptions()
        {
            IncludeHeader = false
        });
    }

    /// <summary>
    /// GetCsvText
    /// </summary>
    /// <param name="dataTable">dataTable</param>
    /// <param name="csvOptions">csvOptions</param>
    /// <returns>csv text</returns>
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
                var val = dataTable.Rows[i][j].ToString()?.Replace(csvOptions.QuoteString, $"{csvOptions.QuoteString}{csvOptions.QuoteString}");
                if (val is { Length: > 0 })
                {
                    data.Append(val.IndexOf(csvOptions.SeparatorCharacter) > -1 ? $"{csvOptions.QuoteString}{val}{csvOptions.QuoteString}" : val);
                }
            }

            data.AppendLine();
        }

        return data.ToString();
    }
}
