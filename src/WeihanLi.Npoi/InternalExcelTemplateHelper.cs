using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Abstract;

namespace WeihanLi.Npoi
{
    internal class InternalExcelTemplateHelper
    {
        public static ISheet EntityListToSheetByTemplate<TEntity>(
            ISheet sheet,
            IEnumerable<TEntity> entityList,
            object? extraData = null)
        {
            var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();
            var propertyColumnDictionary = InternalHelper.GetPropertyColumnDictionary(configuration);

            var globalDictionary = extraData.ParseParamInfo()
                .ToDictionary(x => NpoiTemplateHelper.TemplateOptions.TemplateGlobalParamFormat.FormatWith(x.Key), x => x.Value);
            foreach (var propertyConfiguration in propertyColumnDictionary)
            {
                globalDictionary.Add(NpoiTemplateHelper.TemplateOptions.TemplateHeaderParamFormat.FormatWith(propertyConfiguration.Key.Name), propertyConfiguration.Value.ColumnTitle!);
            }

            var dataFuncDictionary = propertyColumnDictionary
                .ToDictionary(x => NpoiTemplateHelper.TemplateOptions.TemplateDataParamFormat.FormatWith(x.Key.Name), x => x.Key.GetValueGetter<TEntity>());
            foreach (var key in propertyColumnDictionary.Keys)
            {
                if (InternalCache.OutputFormatterFuncCache.TryGetValue(key, out var formatterFunc) && formatterFunc?.Method != null)
                {
                    dataFuncDictionary[NpoiTemplateHelper.TemplateOptions.TemplateDataParamFormat.FormatWith(key.Name)] = entity =>
                    {
                        var val = key.GetValueGetter<TEntity>()?.Invoke(entity);
                        try
                        {
                            var formattedValue = formatterFunc.DynamicInvoke(entity, val);
                            return formattedValue;
                        }
                        catch (Exception e)
                        {
                            Debug.WriteLine(e);
                            InvokeHelper.OnInvokeException?.Invoke(e);
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
                if (row is null)
                {
                    continue;
                }
                for (var cellIndex = row.FirstCellNum; cellIndex < row.LastCellNum; cellIndex++)
                {
                    var cell = row.GetCell(cellIndex);
                    if (cell is null)
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
                                if (cellValue!.Contains(NpoiTemplateHelper.TemplateOptions.TemplateDataEnd))
                                {
                                    dataRowsCount = rowIndex - dataStartRow + 1;
                                    cellValue = cellValue.Replace(NpoiTemplateHelper.TemplateOptions.TemplateDataEnd, string.Empty);
                                }
                            }
                            else
                            {
                                if (cellValue!.Contains(NpoiTemplateHelper.TemplateOptions.TemplateDataBegin))
                                {
                                    dataStartRow = rowIndex;
                                    cellValue = cellValue.Replace(NpoiTemplateHelper.TemplateOptions.TemplateDataBegin, string.Empty);
                                }
                            }
                        }

                        foreach (var param in globalDictionary.Keys)
                        {
                            if (cellValue!.Contains(param))
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
                        var row = sheet.CopyRow(dataStartRow + dataRowsCount + i, dataStartRow + i);
                        for (var j = 0; j < row.LastCellNum; j++)
                        {
                            var cell = row.GetCell(j);
                            if (null != cell)
                            {
                                var cellValue = cell.GetCellValue<string>();
                                if (!string.IsNullOrEmpty(cellValue) && cellValue!.Contains(NpoiTemplateHelper.TemplateOptions.TemplateDataPrefix))
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
