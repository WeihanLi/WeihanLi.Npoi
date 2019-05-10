using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using WeihanLi.Extensions;

namespace WeihanLi.Npoi
{
    /// <summary>
    /// CsvHelper
    /// </summary>
    public static class CsvHelper
    {
        /// <summary>
        /// save to csv file
        /// </summary>
        public static int ToCsvFile(this DataTable dt, string filePath) => ToCsvFile(dt, filePath, true);

        /// <summary>
        /// save to csv file
        /// </summary>
        public static int ToCsvFile(this DataTable dataTable, string filePath, bool includeHeader)
        {
            var csvText = GetCsvText(dataTable, includeHeader);
            if (csvText.IsNullOrWhiteSpace())
            {
                return 0;
            }
            var dir = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            File.WriteAllText(filePath, csvText, Encoding.UTF8);
            return 1;
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
        /// <param name="filePath">csv file path</param>
        public static DataTable ToDataTable(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new ArgumentException(Resource.FileNotFound, nameof(filePath));
            }
            var dt = new DataTable();

            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var sr = new StreamReader(fs, Encoding.UTF8))
                {
                    string strLine;
                    var isFirst = true;
                    while ((strLine = sr.ReadLine()).IsNotNullOrEmpty())
                    {
                        Debug.Assert(strLine != null, nameof(strLine) + " is null");
                        var rowData = strLine.Trim().Split(new[] { InternalConstants.CsvSeparatorCharacter });
                        var dtColumns = rowData.Length;
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
                            for (var j = 0; j < dtColumns; j++)
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
        /// convert csv file data to entity list
        /// </summary>
        /// <param name="filePath">csv file path</param>
        public static List<TEntity> ToEntityList<TEntity>(string filePath) where TEntity : new()
        {
            if (!File.Exists(filePath))
            {
                throw new ArgumentException(Resource.FileNotFound, nameof(filePath));
            }
            var entities = new List<TEntity>();

            if (typeof(TEntity).IsBasicType())
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var sr = new StreamReader(fs, Encoding.UTF8))
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
                            entities.Add(strLine.Trim().To<TEntity>());
                        }
                    }
                }
            }
            else
            {
                IReadOnlyList<PropertyInfo> props = InternalHelper.GetPropertiesForCsvHelper<TEntity>();

                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var sr = new StreamReader(fs, Encoding.UTF8))
                    {
                        string strLine;
                        var isFirstLine = true;
                        while ((strLine = sr.ReadLine()).IsNotNullOrEmpty())
                        {
                            var cols = strLine.Split(new[] { InternalConstants.CsvSeparatorCharacter });
                            if (isFirstLine)
                            {
                                isFirstLine = false;
                            }
                            else
                            {
                                var entity = new TEntity();
                                if (typeof(TEntity).IsValueType)
                                {
                                    var obj = (object)entity;// boxing for value types
                                    for (int i = 0; i < cols.Length; i++)
                                    {
                                        props[i].GetValueSetter().Invoke(obj, cols[i].ToOrDefault(props[i].PropertyType));
                                    }
                                    entity = (TEntity)obj;// unboxing
                                }
                                else
                                {
                                    for (int i = 0; i < cols.Length; i++)
                                    {
                                        props[i].GetValueSetter().Invoke(entity, cols[i].ToOrDefault(props[i].PropertyType));
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
        /// save to csv file
        /// </summary>
        public static int ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath) => ToCsvFile(entities, filePath, true);

        /// <summary>
        /// save to csv file
        /// </summary>
        public static int ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath, bool includeHeader)
        {
            var csvTextData = GetCsvText(entities, includeHeader);
            if (csvTextData.IsNullOrWhiteSpace())
            {
                return 0;
            }
            var dir = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            File.WriteAllText(filePath, csvTextData);

            return 1;
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
                IReadOnlyList<PropertyInfo> props = InternalHelper.GetPropertiesForCsvHelper<TEntity>();
                if (includeHeader)
                {
                    for (var i = 0; i < props.Count; i++)
                    {
                        if (i > 0)
                        {
                            data.Append(InternalConstants.CsvSeparator);
                        }
                        data.Append(props[i].Name);
                    }
                    data.AppendLine();
                }
                foreach (var entity in entities)
                {
                    for (var i = 0; i < props.Count; i++)
                    {
                        if (i > 0)
                        {
                            data.Append(InternalConstants.CsvSeparator);
                        }
                        // https://stackoverflow.com/questions/4617935/is-there-a-way-to-include-commas-in-csv-columns-without-breaking-the-formatting
                        var val = props[i].GetValueGetter().Invoke(entity)?.ToString().Replace("\"", "\"\"");
                        if (!string.IsNullOrEmpty(val))
                        {
                            data.Append(val.IndexOf(',') > -1 ? $"\"{val}\"" : val);
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
                        data.Append(InternalConstants.CsvSeparator);
                    }
                    data.Append(dataTable.Columns[i].ColumnName);
                }
                data.AppendLine();
            }
            for (var i = 0; i < dataTable.Rows.Count; i++)
            {
                for (var j = 0; j < dataTable.Columns.Count; j++)
                {
                    if (i > 0)
                    {
                        data.Append(InternalConstants.CsvSeparator);
                    }
                    // https://stackoverflow.com/questions/4617935/is-there-a-way-to-include-commas-in-csv-columns-without-breaking-the-formatting
                    var val = dataTable.Rows[i][j]?.ToString().Replace("\"", "\"\"");
                    if (!string.IsNullOrEmpty(val))
                    {
                        data.Append(val.IndexOf(',') > -1 ? $"\"{val}\"" : val);
                    }
                }
                data.AppendLine();
            }
            return data.ToString();
        }
    }
}
