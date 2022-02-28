
// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the MIT license.

using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using WeihanLi.Common;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Configurations;
using CellType = WeihanLi.Npoi.Abstract.CellType;
using ICell = WeihanLi.Npoi.Abstract.ICell;

namespace WeihanLi.Npoi;

internal static class InternalExtensions
{
    /// <summary>
    ///     Parse obj to paramDictionary
    /// </summary>
    /// <param name="paramInfo">param object</param>
    /// <returns></returns>
    public static IDictionary<string, object?> ParseParamInfo(this object? paramInfo)
    {
        var paramDic = new Dictionary<string, object?>();
        if (paramInfo is null)
        {
            return paramDic;
        }

        if (paramInfo.IsValueTuple()) // Tuple
        {
            var fields = paramInfo.GetFields();
            foreach (var field in fields)
            {
                paramDic[field.Name] = field.GetValue(paramInfo);
            }
        }
        else if (paramInfo is IDictionary<string, object?> paramDictionary)
        {
            return paramDictionary;
        }
        else // get properties
        {
            var properties = CacheUtil.GetTypeProperties(paramInfo.GetType());
            foreach (var property in properties)
            {
                if (property.CanRead)
                {
                    paramDic[property.Name] = property.GetValueGetter()?.Invoke(paramInfo);
                }
            }
        }

        return paramDic;
    }

    /// <summary>
    ///     GetCellValue
    /// </summary>
    /// <param name="cell">cell</param>
    /// <param name="propertyType">propertyType</param>
    /// <param name="formulaEvaluator">formulaEvaluator</param>
    /// <returns>cellValue</returns>
    public static object? GetCellValue(this ICell? cell, Type propertyType, IFormulaEvaluator? formulaEvaluator)
    {
        if (cell is null || cell.CellType == CellType.Blank || cell.CellType == CellType.Error)
        {
            return propertyType.GetDefaultValue();
        }

        return cell.Value?.ToOrDefault(propertyType);
    }

    /// <summary>
    ///     GetCellValue
    /// </summary>
    /// <typeparam name="T">Type</typeparam>
    /// <param name="cell">cell</param>
    /// <returns></returns>
    public static T? GetCellValue<T>(this ICell? cell) => (cell?.Value).ToOrDefault<T>();

    /// <summary>
    ///     SetCellValue
    /// </summary>
    /// <param name="cell">ICell</param>
    /// <param name="value">value</param>
    public static void SetCellValue(this ICell? cell, object? value) => cell?.SetCellValue(value, null);

    /// <summary>
    ///     SetCellValue
    /// </summary>
    /// <param name="cell">ICell</param>
    /// <param name="value">value</param>
    /// <param name="formatter">formatter</param>
    public static void SetCellValue(this ICell cell, object? value, string? formatter)
    {
        if (value is null)
        {
            cell.CellType = CellType.Blank;
            return;
        }

        if (value is DateTime time)
        {
            cell.Value = string.IsNullOrWhiteSpace(formatter)
                ? time.Date == time ? time.ToStandardDateString() : time.ToStandardTimeString()
                : time.ToString(formatter);
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
                cell.Value = value is IFormattable val && formatter.IsNotNullOrWhiteSpace()
                    ? val.ToString(formatter, CultureInfo.CurrentCulture)
                    : value.ToString();
                cell.CellType = CellType.String;
            }
        }
    }

    /// <summary>
    ///     GetPropertySettingByPropertyName
    /// </summary>
    /// <param name="mappingDictionary">mappingDictionary</param>
    /// <param name="propertyName">propertyName</param>
    /// <returns></returns>
    internal static PropertyConfiguration? GetPropertySettingByPropertyName(
        this IDictionary<PropertyInfo, PropertyConfiguration> mappingDictionary, string propertyName)
        => mappingDictionary.Values.FirstOrDefault(_ => _.PropertyName.EqualsIgnoreCase(propertyName));

    /// <summary>
    ///     GetPropertyConfigurationByColumnName
    /// </summary>
    /// <param name="mappingDictionary">mappingDictionary</param>
    /// <param name="columnTitle">columnTitle</param>
    /// <returns></returns>
    internal static PropertyConfiguration? GetPropertySetting(
        this IDictionary<PropertyInfo, PropertyConfiguration> mappingDictionary, string columnTitle) =>
        mappingDictionary.Values.FirstOrDefault(k => k.ColumnTitle.EqualsIgnoreCase(columnTitle)) ??
        mappingDictionary.GetPropertySettingByPropertyName(columnTitle);
}
