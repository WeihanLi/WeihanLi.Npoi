using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi
{
    /// <summary>
    /// CsvHelper
    /// </summary>
    public static class CsvHelper
    {
        public static char CsvSeparatorCharacter = ',';

        /// <summary>
        /// save to csv file
        /// </summary>
        public static bool ToCsvFile(this DataTable dt, string filePath) => ToCsvFile(dt, filePath, true);

        /// <summary>
        /// save to csv file
        /// </summary>
        public static bool ToCsvFile(this DataTable dataTable, string filePath, bool includeHeader)
        {
            var dir = Path.GetDirectoryName(filePath);
            if (dir == null)
            {
                throw new ArgumentException(Resource.InvalidFilePath);
            }
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var csvText = GetCsvText(dataTable, includeHeader);
            if (csvText.IsNullOrWhiteSpace())
            {
                return false;
            }

            File.WriteAllText(filePath, csvText, Encoding.UTF8);
            return true;
        }

        /// <summary>
        /// to csv bytes
        /// </summary>
        public static byte[] ToCsvBytes(this DataTable dt) => ToCsvBytes(dt, true);

        /// <summary>
        /// to csv bytes
        /// </summary>
        public static byte[] ToCsvBytes(this DataTable dataTable, bool includeHeader) => GetCsvText(dataTable, includeHeader).GetBytes();

        /// <summary>
        /// convert csv file data to dataTable
        /// </summary>
        /// <param name="csvBytes">csv bytes</param>
        public static DataTable ToDataTable(byte[] csvBytes)
        {
            if (null == csvBytes)
            {
                throw new ArgumentNullException(nameof(csvBytes));
            }

            using (var ms = new MemoryStream(csvBytes))
            {
                return ToDataTable(ms);
            }
        }

        /// <summary>
        /// convert csv stream data to dataTable
        /// </summary>
        /// <param name="stream">stream</param>
        public static DataTable ToDataTable(Stream stream)
        {
            if (null == stream)
            {
                throw new ArgumentNullException(nameof(stream));
            }
            var dt = new DataTable();

            if (stream.CanRead)
            {
                using (var sr = new StreamReader(stream, Encoding.UTF8))
                {
                    string strLine;
                    var isFirst = true;
                    while ((strLine = sr.ReadLine()).IsNotNullOrEmpty())
                    {
                        Debug.Assert(strLine != null, nameof(strLine) + " is null");

                        var rowData = ParseLine(strLine);
                        var dtColumns = rowData.Count;
                        if (isFirst)
                        {
                            for (var i = 0; i < dtColumns; i++)
                            {
                                dt.Columns.Add(rowData[i]);
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
            }

            return dt;
        }

        /// <summary>
        /// convert csv file data to dataTable
        /// </summary>
        /// <param name="filePath">csv file path</param>
        public static DataTable ToDataTable(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new ArgumentException(Resource.FileNotFound, nameof(filePath));
            }

            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                return ToDataTable(fs);
            }
        }

        /// <summary>
        /// convert csv file data to entity list
        /// </summary>
        /// <param name="filePath">csv file path</param>
        public static List<TEntity> ToEntityList<TEntity>(string filePath) where TEntity : new()
        {
            if (!File.Exists(filePath))
            {
                throw new ArgumentException(Resource.FileNotFound, nameof(filePath));
            }

            return ToEntityList<TEntity>(File.ReadAllBytes(filePath));
        }

        /// <summary>
        /// convert csv file data to entity list
        /// </summary>
        /// <param name="csvBytes">csv bytes</param>
        public static List<TEntity> ToEntityList<TEntity>(byte[] csvBytes) where TEntity : new()
        {
            if (null == csvBytes)
            {
                throw new ArgumentNullException(nameof(csvBytes));
            }
            var entities = new List<TEntity>();

            if (typeof(TEntity).IsBasicType())
            {
                using (var ms = new MemoryStream(csvBytes))
                {
                    using (var sr = new StreamReader(ms, Encoding.UTF8))
                    {
                        string strLine;
                        var isFirstLine = true;
                        while ((strLine = sr.ReadLine()).IsNotNullOrEmpty())
                        {
                            if (isFirstLine)
                            {
                                isFirstLine = false;
                                continue;
                            }
                            //
                            Debug.Assert(strLine != null, nameof(strLine) + " is null");
                            entities.Add(strLine.Trim().To<TEntity>());
                        }
                    }
                }
            }
            else
            {
                var propertyColumnDictionary = InternalHelper.GetPropertyColumnDictionary<TEntity>();
                var propertyColumnDic = propertyColumnDictionary.ToDictionary(_ => _.Key, _ => new PropertyConfiguration()
                {
                    ColumnIndex = -1,
                    ColumnFormatter = _.Value.ColumnFormatter,
                    ColumnTitle = _.Value.ColumnTitle,
                    ColumnWidth = _.Value.ColumnWidth,
                    IsIgnored = _.Value.IsIgnored
                });
                using (var ms = new MemoryStream(csvBytes))
                {
                    using (var sr = new StreamReader(ms, Encoding.UTF8))
                    {
                        string strLine;
                        var isFirstLine = true;
                        while ((strLine = sr.ReadLine()).IsNotNullOrEmpty())
                        {
                            var entityType = typeof(TEntity);
                            var cols = ParseLine(strLine);
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
                                continue;
                            }
                            else
                            {
                                var entity = new TEntity();
                                if (entityType.IsValueType)
                                {
                                    var obj = (object)entity;// boxing for value types

                                    foreach (var key in propertyColumnDic.Keys)
                                    {
                                        var colIndex = propertyColumnDic[key].ColumnIndex;
                                        if (colIndex >= 0 && colIndex < cols.Count && key.CanWrite)
                                        {
                                            var columnValue = key.PropertyType.GetDefaultValue();
                                            var valueApplied = false;
                                            if (InternalCache.ColumnInputFormatterFuncCache.TryGetValue(key, out var formatterFunc) && formatterFunc?.Method != null)
                                            {
                                                var cellValue = cols[colIndex];
                                                try
                                                {
                                                    // apply custom formatterFunc
                                                    columnValue = formatterFunc.Invoke(new object[] { cellValue });
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

                                            key.GetValueSetter()?.Invoke(entity, columnValue);
                                        }
                                    }

                                    entity = (TEntity)obj;// unboxing
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
                                            if (InternalCache.ColumnInputFormatterFuncCache.TryGetValue(key, out var formatterFunc) && formatterFunc?.Method != null)
                                            {
                                                var cellValue = cols[colIndex];
                                                try
                                                {
                                                    // apply custom formatterFunc
                                                    columnValue = formatterFunc.Invoke(new object[] { cellValue });
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

                                            key.GetValueSetter()?.Invoke(entity, columnValue);
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
                                            if (InternalCache.InputFormatterFuncCache.TryGetValue(propertyInfo, out var formatterFunc) && formatterFunc?.Method != null)
                                            {
                                                try
                                                {
                                                    // apply custom formatterFunc
                                                    var formattedValue = formatterFunc.Invoke(new[] { entity, propertyValue });
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

                                entities.Add(entity);
                            }
                        }
                    }
                }
            }

            return entities;
        }

        /// <summary>
        /// convert csv file data to entity list
        /// </summary>
        /// <param name="csvStream">csv Stream</param>
        public static List<TEntity> ToEntityList<TEntity>(Stream csvStream) where TEntity : new()
        {
            if (null == csvStream)
            {
                throw new ArgumentNullException(nameof(csvStream));
            }
            return ToEntityList<TEntity>(csvStream.ToByteArray());
        }

        private static IReadOnlyList<string> ParseLine(string line)
        {
            if (string.IsNullOrEmpty(line))
                return new[] { string.Empty };

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
                    if (character == '"')
                    {
                        inQuotes = true;
                        continue;
                    }
                }

                // If we are in between double quotes
                if (inQuotes)
                {
                    if ((i + 1) == line.Length)
                    {
                        fields.Add(line[i].ToString());
                        break;
                    }

                    if (character == '"' && line[i + 1] == CsvSeparatorCharacter) // quotes end
                    {
                        inQuotes = false;
                        inColumn = false;
                        i++; //skip next
                    }
                    else if (character == '"' && line[i + 1] == '"') // quotes
                    {
                        i++; //skip next
                    }
                    else if (character == '"')
                    {
                        throw new ArgumentException($"unable to escape {line}");
                    }
                }
                else if (character == CsvSeparatorCharacter)
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
        /// save to csv file
        /// </summary>
        public static bool ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath) => ToCsvFile(entities, filePath, true);

        /// <summary>
        /// save to csv file
        /// </summary>
        public static bool ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath, bool includeHeader)
        {
            var dir = Path.GetDirectoryName(filePath);
            if (dir == null)
            {
                throw new ArgumentException(Resource.InvalidFilePath);
            }
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var csvTextData = GetCsvText(entities, includeHeader);
            if (csvTextData.IsNullOrWhiteSpace())
            {
                return false;
            }

            File.WriteAllText(filePath, csvTextData, Encoding.UTF8);
            return true;
        }

        /// <summary>
        /// to csv bytes
        /// </summary>
        public static byte[] ToCsvBytes<TEntity>(this IEnumerable<TEntity> entities) => ToCsvBytes(entities, true);

        /// <summary>
        /// to csv bytes
        /// </summary>
        public static byte[] ToCsvBytes<TEntity>(this IEnumerable<TEntity> entities, bool includeHeader) => GetCsvText(entities, includeHeader).GetBytes();

        private static string GetCsvText<TEntity>(this IEnumerable<TEntity> entities, bool includeHeader)
        {
            if (entities == null)
            {
                return string.Empty;
            }

            var data = new StringBuilder();
            var isBasicType = typeof(TEntity).IsBasicType();
            if (isBasicType)
            {
                if (includeHeader)
                {
                    data.AppendLine(InternalConstants.DefaultPropertyNameForBasicType);
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
                if (includeHeader)
                {
                    for (var i = 0; i < props.Count; i++)
                    {
                        if (i > 0)
                        {
                            data.Append(CsvSeparatorCharacter);
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
                        if (InternalCache.OutputFormatterFuncCache.TryGetValue(props[i], out var formatterFunc) && formatterFunc?.Method != null)
                        {
                            try
                            {
                                // apply custom formatterFunc
                                propertyValue = formatterFunc.Invoke(new[] { entity, propertyValue });
                            }
                            catch (Exception e)
                            {
                                Debug.WriteLine(e);
                                InvokeHelper.OnInvokeException?.Invoke(e);
                            }
                        }

                        if (i > 0)
                        {
                            data.Append(CsvSeparatorCharacter);
                        }
                        // https://stackoverflow.com/questions/4617935/is-there-a-way-to-include-commas-in-csv-columns-without-breaking-the-formatting
                        var val = propertyValue?.ToString().Replace("\"", "\"\"");
                        if (!string.IsNullOrEmpty(val))
                        {
                            data.Append(val.IndexOf(CsvSeparatorCharacter) > -1 ? $"\"{val}\"" : val);
                        }
                    }
                    data.AppendLine();
                }
            }

            return data.ToString();
        }

        private static string GetCsvText(this DataTable dataTable, bool includeHeader)
        {
            if (dataTable == null || dataTable.Rows.Count == 0 || dataTable.Columns.Count == 0)
            {
                return string.Empty;
            }

            var data = new StringBuilder();

            if (includeHeader)
            {
                for (var i = 0; i < dataTable.Columns.Count; i++)
                {
                    if (i > 0)
                    {
                        data.Append(CsvSeparatorCharacter);
                    }
                    data.Append(dataTable.Columns[i].ColumnName);
                }
                data.AppendLine();
            }
            for (var i = 0; i < dataTable.Rows.Count; i++)
            {
                for (var j = 0; j < dataTable.Columns.Count; j++)
                {
                    if (j > 0)
                    {
                        data.Append(CsvSeparatorCharacter);
                    }
                    // https://stackoverflow.com/questions/4617935/is-there-a-way-to-include-commas-in-csv-columns-without-breaking-the-formatting
                    var val = dataTable.Rows[i][j]?.ToString().Replace("\"", "\"\"");
                    if (!string.IsNullOrEmpty(val))
                    {
                        data.Append(val.IndexOf(CsvSeparatorCharacter) > -1 ? $"\"{val}\"" : val);
                    }
                }
                data.AppendLine();
            }
            return data.ToString();
        }
    }
}
