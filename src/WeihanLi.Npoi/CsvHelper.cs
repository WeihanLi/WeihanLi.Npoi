using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
                        var rowData = strLine.Trim().Split(new[] { InternalConstants.CsvSeparator }, StringSplitOptions.RemoveEmptyEntries);
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
                var propertyColumnDictionary = InternalHelper.GetExcelConfigurationMapping<TEntity>().PropertyConfigurationDictionary.Where(_ => !_.Value.PropertySetting.IsIgnored).ToDictionary(_ => _.Key, _ => _.Value.PropertySetting);

                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var sr = new StreamReader(fs, Encoding.UTF8))
                    {
                        string strLine;
                        var isFirstLine = true;
                        while ((strLine = sr.ReadLine()).IsNotNullOrEmpty())
                        {
                            var cols = strLine.Split(new[] { InternalConstants.CsvSeparatorCharactor }, StringSplitOptions.RemoveEmptyEntries);
                            if (isFirstLine)
                            {
                                for (int i = 0; i < cols.Length; i++)
                                {
                                    var col = propertyColumnDictionary.GetPropertySetting(cols[i].Trim());
                                    if (null != col)
                                    {
                                        col.ColumnIndex = i;
                                    }
                                }
                                isFirstLine = false;
                            }
                            else
                            {
                                var entity = new TEntity();
                                if (typeof(TEntity).IsValueType)
                                {
                                    var obj = (object)entity;// boxing for value types
                                    foreach (var key in propertyColumnDictionary.Keys)
                                    {
                                        var colIndex = propertyColumnDictionary[key].ColumnIndex;
                                        key.GetValueSetter().Invoke(obj, cols[colIndex]);
                                    }
                                    entity = (TEntity)obj;// unboxing
                                }
                                else
                                {
                                    foreach (var key in propertyColumnDictionary.Keys)
                                    {
                                        var colIndex = propertyColumnDictionary[key].ColumnIndex;
                                        key.GetValueSetter().Invoke(entity, cols[colIndex]);
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
            if (entities == null || !entities.Any())
            {
                return string.Empty;
            }
            var isBasicType = typeof(TEntity).IsBasicType();
            IReadOnlyList<PropertyInfo> props = InternalHelper.GetPropertiesForCsvHelper<TEntity>();
            if (!isBasicType && props.Count == 0)
            {
                return string.Empty;
            }
            var data = new StringBuilder();
            if (includeHeader)
            {
                if (isBasicType)
                {
                    data.Append(InternalConstants.DefaultPropertyNameForBasicType);
                }
                else
                {
                    for (var i = 0; i < props.Count; i++)
                    {
                        if (i > 0)
                        {
                            data.Append(InternalConstants.CsvSeparator);
                        }
                        data.Append(props[i].Name);
                    }
                }
                data.AppendLine();
            }

            if (isBasicType)
            {
                foreach (var entity in entities)
                {
                    data.AppendLine(Convert.ToString(entity));
                }
            }
            else
            {
                foreach (var entity in entities)
                {
                    for (var i = 0; i < props.Count; i++)
                    {
                        if (i > 0)
                        {
                            data.Append(InternalConstants.CsvSeparator);
                        }
                        data.Append(props[i].GetValueGetter().Invoke(entity));
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
                    data.Append(dataTable.Rows[i][j]);
                }
                data.AppendLine();
            }
            return data.ToString();
        }
    }
}
