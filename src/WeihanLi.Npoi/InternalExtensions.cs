using JetBrains.Annotations;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Configurations;
using CellType = WeihanLi.Npoi.Abstract.CellType;
using ICell = WeihanLi.Npoi.Abstract.ICell;

namespace WeihanLi.Npoi
{
    internal static class InternalExtensions
    {
        /// <summary>
        ///     GetCellValue
        /// </summary>
        /// <param name="cell">cell</param>
        /// <param name="propertyType">propertyType</param>
        /// <returns>cellValue</returns>
        public static object GetCellValue([CanBeNull] this ICell cell, Type propertyType)
        {
            if (cell == null || cell.CellType == CellType.Blank || cell.CellType == CellType.Error)
            {
                return propertyType.GetDefaultValue();
            }
            return cell.Value.ToOrDefault(propertyType);
        }

        /// <summary>
        ///     GetCellValue
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="cell">cell</param>
        /// <returns></returns>
        public static T GetCellValue<T>([CanBeNull] this ICell cell) => (cell?.Value).ToOrDefault<T>();

        /// <summary>
        ///     SetCellValue
        /// </summary>
        /// <param name="cell">ICell</param>
        /// <param name="value">value</param>
        public static void SetCellValue([NotNull] this ICell cell, object value) => cell.SetCellValue(value, null);

        /// <summary>
        ///     SetCellValue
        /// </summary>
        /// <param name="cell">ICell</param>
        /// <param name="value">value</param>
        /// <param name="formatter">formatter</param>
        public static void SetCellValue([NotNull] this ICell cell, object value, string formatter)
        {
            if (null == value)
            {
                cell.CellType = CellType.Blank;
                return;
            }

            if (value is DateTime time)
            {
                cell.Value = (string.IsNullOrWhiteSpace(formatter)
                    ? (time.Date == time ? time.ToStandardDateString() : time.ToStandardTimeString())
                    : time.ToString(formatter));
                cell.CellType = CellType.String;
            }
            else
            {
                var type = value.GetType();
                if (
                    type == typeof(double) ||
                    type == typeof(int) ||
                    type == typeof(long) ||
                    type == typeof(float) ||
                    type == typeof(decimal)
                )
                {
                    cell.Value = Convert.ToDouble(value);
                    cell.CellType = CellType.Numeric;
                }
                else if (type == typeof(bool))
                {
                    cell.Value = value;
                    cell.CellType = CellType.Boolean;
                }
                else
                {
                    cell.Value = (value is IFormattable val && formatter.IsNotNullOrWhiteSpace()
                        ? val.ToString(formatter, CultureInfo.CurrentCulture)
                        : value.ToString());
                    cell.CellType = CellType.String;
                }
            }
        }

        /// <summary>
        /// GetPropertySettingByPropertyName
        /// </summary>
        /// <param name="mappingDictionary">mappingDictionary</param>
        /// <param name="propertyName">propertyName</param>
        /// <returns></returns>
        internal static PropertyConfiguration GetPropertySettingByPropertyName([NotNull] this IDictionary<PropertyInfo, PropertyConfiguration> mappingDictionary, [NotNull] string propertyName)
            => mappingDictionary.Values.FirstOrDefault(_ => _.PropertyName.EqualsIgnoreCase(propertyName));

        /// <summary>
        /// GetPropertyConfigurationByColumnName
        /// </summary>
        /// <param name="mappingDictionary">mappingDictionary</param>
        /// <param name="columnTitle">columnTitle</param>
        /// <returns></returns>
        internal static PropertyConfiguration GetPropertySetting([NotNull]this IDictionary<PropertyInfo, PropertyConfiguration> mappingDictionary, [NotNull]string columnTitle)
        {
            return mappingDictionary.Values.FirstOrDefault(k => k.ColumnTitle.EqualsIgnoreCase(columnTitle)) ?? mappingDictionary.GetPropertySettingByPropertyName(columnTitle);
        }
    }
}
