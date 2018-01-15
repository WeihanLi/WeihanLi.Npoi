using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using JetBrains.Annotations;
using NPOI.SS.UserModel;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Attributes;

namespace WeihanLi.Npoi
{
    public static class NpoiExtension
    {
        /// <summary>
        /// SetCellValue
        /// </summary>
        /// <param name="cell">ICell</param>
        /// <param name="value">value</param>
        /// <param name="valueType">valueType</param>
        public static void SetCellValue([NotNull]this ICell cell, object value, Type valueType)
        {
            if (null == value)
            {
                cell.SetCellValue("");
                return;
            }
            var type = value.GetType();
            if (type == typeof(DateTime))
            {
                cell.SetCellValue((DateTime)value);
            }
            else if (type == typeof(double) ||
                type == typeof(int) ||
                type == typeof(long) ||
                type == typeof(float) ||
                type == typeof(decimal) ||
                type == typeof(double?) ||
                type == typeof(int?) ||
                type == typeof(long?) ||
                type == typeof(float?) ||
                type == typeof(decimal?) ||
                type == typeof(DateTime?)
            )
            {
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue((double)value);
            }
            else if (type == typeof(bool))
            {
                cell.SetCellType(CellType.Boolean);
                cell.SetCellValue((bool)value);
            }
            else
            {
                cell.SetCellType(CellType.String);
                cell.SetCellValue(value.ToString());
            }
        }

        /// <summary>
        /// GetCellValue
        /// </summary>
        /// <param name="cell">cell</param>
        /// <param name="propertyType">propertyType</param>
        /// <returns>cellValue</returns>
        public static object GetCellValue([NotNull]this ICell cell, Type propertyType)
        {
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue;
                    }
                    return cell.NumericCellValue.To(propertyType);

                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Boolean:
                    return cell.BooleanCellValue;

                default:
                    return cell.ToString().To(propertyType);
            }
        }

        public static void WriteToFile([NotNull] this IWorkbook workbook, string filePath)
        {
            using (var fileStream = File.Create(filePath))
            {
                workbook.Write(fileStream);
            }
        }

        #region Mapping

        internal static PropertyInfo GetPropertyInfo([NotNull]this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, int index) => mappingDictionary.Keys.ToArray()[index];

        internal static PropertyInfo GetPropertyInfo([NotNull]this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, [NotNull]string propertyName) => mappingDictionary.Keys.FirstOrDefault(k => k.Name.EqualsIgnoreCase(propertyName));

        internal static ColumnAttribute GetColumnAttribute([NotNull]this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, int index) => mappingDictionary.Values.ToArray()[index];

        internal static ColumnAttribute GetColumnAttributeByPropertyName(
            [NotNull] this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, [NotNull] string propertyName)
            => mappingDictionary.Keys.Any(_ => _.Name.EqualsIgnoreCase(propertyName)) ?
            mappingDictionary[mappingDictionary.Keys.First(_ => _.Name.EqualsIgnoreCase(propertyName))] : null;

        internal static ColumnAttribute GetColumnAttribute([NotNull]this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, [NotNull]string columnTitle) => mappingDictionary.Values.FirstOrDefault(k => k.Title.EqualsIgnoreCase(columnTitle));

        #endregion Mapping
    }
}