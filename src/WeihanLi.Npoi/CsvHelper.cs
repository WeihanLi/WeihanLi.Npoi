using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
        public static void ToCsvFile(this DataTable dt, string filePath) => ToCsvFile(dt, filePath, true);

        /// <summary>
        /// save to csv file
        /// </summary>
        public static void ToCsvFile(this DataTable dt, string filePath, bool includeHeader)
        {
            var data = new StringBuilder();

            if (includeHeader)
            {
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    data.Append(dt.Columns[i].ColumnName);
                    if (i < dt.Columns.Count - 1)
                    {
                        data.Append(InternalConstants.CsvSeparator);
                    }
                }
                data.AppendLine();
            }
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                for (var j = 0; j < dt.Columns.Count; j++)
                {
                    data.Append(dt.Rows[i][j].ToString());
                    if (j < dt.Columns.Count - 1)
                    {
                        data.Append(InternalConstants.CsvSeparator);
                    }
                }
                data.AppendLine();
            }
            var dir = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            File.WriteAllText(filePath, data.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// to csv bytes
        /// </summary>
        public static byte[] ToCsvBytes(this DataTable dt) => ToCsvBytes(dt, true);

        /// <summary>
        /// to csv bytes
        /// </summary>
        public static byte[] ToCsvBytes(this DataTable dt, bool includeHeader)
        {
            var data = new StringBuilder();

            if (includeHeader)
            {
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    data.Append(dt.Columns[i].ColumnName);
                    if (i < dt.Columns.Count - 1)
                    {
                        data.Append(InternalConstants.CsvSeparator);
                    }
                }
                data.AppendLine();
            }
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                for (var j = 0; j < dt.Columns.Count; j++)
                {
                    data.Append(dt.Rows[i][j].ToString());
                    if (j < dt.Columns.Count - 1)
                    {
                        data.Append(InternalConstants.CsvSeparator);
                    }
                }
                data.AppendLine();
            }

            return data.ToString().GetBytes();
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

            var dt = ToDataTable(filePath);

            return dt.ToEntities<TEntity>().ToList();
        }

        /// <summary>
        /// save to csv file
        /// </summary>
        public static void ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath) => ToCsvFile(entities, filePath, true);

        /// <summary>
        /// save to csv file
        /// </summary>
        public static void ToCsvFile<TEntity>(this IEnumerable<TEntity> entities, string filePath, bool includeHeader)
        {
            var dt = entities.ToDataTable();
            var data = new StringBuilder();

            if (includeHeader)
            {
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    data.Append(dt.Columns[i].ColumnName);
                    if (i < dt.Columns.Count - 1)
                    {
                        data.Append(InternalConstants.CsvSeparator);
                    }
                }
                data.AppendLine();
            }
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                for (var j = 0; j < dt.Columns.Count; j++)
                {
                    data.Append(dt.Rows[i][j].ToString());
                    if (j < dt.Columns.Count - 1)
                    {
                        data.Append(InternalConstants.CsvSeparator);
                    }
                }
                data.AppendLine();
            }

            var dir = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            File.WriteAllText(filePath, data.ToString());
        }

        /// <summary>
        /// to csv bytes
        /// </summary>
        public static byte[] ToCsvBytes<TEntity>(this IEnumerable<TEntity> entities) => ToCsvBytes(entities, true);

        /// <summary>
        /// to csv bytes
        /// </summary>
        public static byte[] ToCsvBytes<TEntity>(this IEnumerable<TEntity> entities, bool includeHeader)
        {
            var dt = entities.ToDataTable();

            var data = new StringBuilder();

            if (includeHeader)
            {
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    data.Append(dt.Columns[i].ColumnName);
                    if (i < dt.Columns.Count - 1)
                    {
                        data.Append(InternalConstants.CsvSeparator);
                    }
                }
                data.AppendLine();
            }
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                for (var j = 0; j < dt.Columns.Count; j++)
                {
                    data.Append(dt.Rows[i][j].ToString());
                    if (j < dt.Columns.Count - 1)
                    {
                        data.Append(InternalConstants.CsvSeparator);
                    }
                }
                data.AppendLine();
            }

            return data.ToString().GetBytes();
        }
    }
}
