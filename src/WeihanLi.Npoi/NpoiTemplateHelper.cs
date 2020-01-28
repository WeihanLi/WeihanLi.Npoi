using JetBrains.Annotations;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;

namespace WeihanLi.Npoi
{
    internal static class NpoiTemplateHelper
    {
        public static readonly TemplateOptions TemplateOptions = new TemplateOptions();

        // export via template
        public static ISheet EntityListToSheetByTemplate<TEntity>(
            [NotNull]ISheet sheet,
            IEnumerable<TEntity> entityList,
            object extraData = null)
        {
            if (null == entityList)
            {
                return sheet;
            }
            var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();
            var propertyColumnDictionary = InternalHelper.GetPropertyColumnDictionary(configuration);

            var globalDictionary = extraData.ParseParamInfo()
                .ToDictionary(x => TemplateOptions.TemplateGlobalParamFormat.FormatWith(x.Key), x => x.Value);
            foreach (var propertyConfiguration in propertyColumnDictionary)
            {
                globalDictionary.Add(TemplateOptions.TemplateHeaderParamFormat.FormatWith(propertyConfiguration.Key.Name), propertyConfiguration.Value.ColumnTitle);
            }

            var dataFuncDictionary = propertyColumnDictionary
                .ToDictionary(x => TemplateOptions.TemplateDataParamFormat.FormatWith(x.Key.Name), x => x.Key.GetValueGetter<TEntity>());
            foreach (var key in propertyColumnDictionary.Keys)
            {
                if (InternalCache.OutputFormatterFuncCache.TryGetValue(key, out var formatterFunc) && formatterFunc != null)
                {
                    dataFuncDictionary[TemplateOptions.TemplateDataParamFormat.FormatWith(key.Name)] = entity =>
                    {
                        var val = key.GetValueGetter<TEntity>()?.Invoke(entity);
                        var funcType = typeof(Func<,,>).MakeGenericType(configuration.EntityType, key.PropertyType, typeof(object));
                        var method = funcType.GetProperty("Method")?.GetValueGetter()?.Invoke(formatterFunc) as MethodInfo;
                        var target = funcType.GetProperty("Target")?.GetValueGetter()?.Invoke(formatterFunc);
                        if (null != method && null != target)
                        {
                            try
                            {
                                var formattedValue = method.Invoke(target, new[] { entity, val });
                                return formattedValue;
                            }
                            catch (Exception e)
                            {
                                Debug.WriteLine(e);
                                InvokeHelper.OnInvokeException?.Invoke(e);
                            }
                        }
                        return val;
                    };
                }
            }

            // parseTemplate
            int dataStartRow = -1, dataRowsCount = 0;
            for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    continue;
                }
                for (var cellIndex = row.FirstCellNum; cellIndex < row.LastCellNum; cellIndex++)
                {
                    var cell = row.GetCell(cellIndex);
                    if (cell == null)
                    {
                        continue;
                    }

                    var cellValue = cell.GetCellValue<string>();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        var beforeValue = cellValue;
                        if (dataStartRow <= 0 || dataRowsCount <= 0)
                        {
                            if (dataStartRow >= 0)
                            {
                                if (cellValue.Contains(TemplateOptions.TemplateDataEnd))
                                {
                                    dataRowsCount = rowIndex - dataStartRow + 1;
                                    cellValue = cellValue.Replace(TemplateOptions.TemplateDataEnd, string.Empty);
                                }
                            }
                            else
                            {
                                if (cellValue.Contains(TemplateOptions.TemplateDataBegin))
                                {
                                    dataStartRow = rowIndex;
                                    cellValue = cellValue.Replace(TemplateOptions.TemplateDataBegin, string.Empty);
                                }
                            }
                        }

                        foreach (var param in globalDictionary.Keys)
                        {
                            if (cellValue.Contains(param))
                            {
                                cellValue = cellValue
                                    .Replace(param,
                                    globalDictionary[param]?.ToString() ?? string.Empty);
                            }
                        }

                        if (beforeValue != cellValue)
                        {
                            cell.SetCellValue(cellValue);
                        }
                    }
                }
            }

            if (dataStartRow >= 0 && dataRowsCount > 0)
            {
                foreach (var entity in entityList)
                {
                    sheet.ShiftRows(dataStartRow, sheet.LastRowNum, dataRowsCount);
                    for (var i = 0; i < dataRowsCount; i++)
                    {
                        sheet.CopyRow(dataStartRow + dataRowsCount + i, dataStartRow + i);
                        var row = sheet.GetRow(dataStartRow + i);
                        if (null != row)
                        {
                            for (var j = 0; j < row.LastCellNum; j++)
                            {
                                var cell = row.GetCell(j);
                                if (null != cell)
                                {
                                    var cellValue = cell.GetCellValue<string>();
                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains(TemplateOptions.TemplateDataPrefix))
                                    {
                                        var beforeValue = cellValue;

                                        foreach (var param in dataFuncDictionary.Keys)
                                        {
                                            if (cellValue.Contains(param))
                                            {
                                                cellValue = cellValue.Replace(param,
                                                    dataFuncDictionary[param]?.Invoke(entity)?.ToString() ?? string.Empty);
                                            }
                                        }

                                        if (beforeValue != cellValue)
                                        {
                                            cell.SetCellValue(cellValue);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //
                    dataStartRow += dataRowsCount;
                }

                // remove data template
                for (var i = 0; i < dataRowsCount; i++)
                {
                    var row = sheet.GetRow(dataStartRow + i);
                    if (null != row)
                    {
                        sheet.RemoveRow(row);
                    }
                }
                sheet.ShiftRows(dataStartRow + dataRowsCount, sheet.LastRowNum, -dataRowsCount);
            }

            return sheet;
        }
    }
}
