// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Data;
using System.Diagnostics;
using System.Reflection;
using WeihanLi.Common;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Configurations;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi;

internal static class NpoiHelper
{
    private static SheetSetting GetSheetSetting(IDictionary<int, SheetSetting> sheetSettings, int sheetIndex) =>
        sheetIndex > 0 && sheetSettings.ContainsKey(sheetIndex)
            ? sheetSettings[sheetIndex]
            : sheetSettings[0];

    public static IEnumerable<TEntity?> SheetToEntities<TEntity>(ISheet? sheet, int sheetIndex, Action<TEntity?, ExcelConfiguration<TEntity>, int>? dataAction = null) where TEntity : new()
    {
        if (sheet is null || sheet.PhysicalNumberOfRows <= 0)
        {
            yield break;
        }

        var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();
        var sheetSetting = GetSheetSetting(configuration.SheetSettings, sheetIndex);

        var propertyColumnDictionary = InternalHelper.GetPropertyColumnDictionary(configuration);
        var propertyColumnDic = sheetSetting.HeaderRowIndex >= 0
            ? propertyColumnDictionary.ToDictionary(_ => _.Key,
                _ => new PropertyConfiguration
                {
                    ColumnIndex = -1,
                    ColumnFormatter = _.Value.ColumnFormatter,
                    ColumnTitle = _.Value.ColumnTitle,
                    ColumnWidth = _.Value.ColumnWidth,
                    IsIgnored = _.Value.IsIgnored
                })
            : propertyColumnDictionary;
        var formulaEvaluator = sheet.Workbook.GetFormulaEvaluator();

        var pictures = propertyColumnDic
            .Any(p => p.Key.CanWrite &&
                      (p.Key.PropertyType == typeof(byte[]) || p.Key.PropertyType == typeof(IPictureData)))
            ? sheet.GetPicturesAndPosition()
            : new Dictionary<CellPosition, IPictureData>();

        for (var rowIndex = sheet.FirstRowNum;
            rowIndex <= (sheetSetting.EndRowIndex ?? sheet.LastRowNum);
            rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);

            if (rowIndex == sheetSetting.HeaderRowIndex) // readerHeader
            {
                if (row != null)
                {
                    // adjust column index according to the imported data header
                    for (var i = row.FirstCellNum; i < row.LastCellNum; i++)
                    {
                        if (row.GetCell(i) is null)
                        {
                            continue;
                        }

                        var title = row.GetCell(i).StringCellValue.Trim();
                        var col = propertyColumnDic.GetPropertySetting(title);
                        if (null != col)
                        {
                            col.ColumnIndex = i;
                        }
                    }
                }

                // use default column index if no headers
                if (propertyColumnDic.Values.Any(_ => _.ColumnIndex < 0))
                {
                    propertyColumnDic = propertyColumnDictionary;
                }
            }
            else if (rowIndex >= sheetSetting.StartRowIndex)
            {
                if (sheetSetting.RowFilter?.Invoke(row) == false)
                {
                    continue;
                }

                if (row is null)
                {
                    yield return default;
                }
                else
                {
                    TEntity? entity;
                    if (row.Cells.Count > 0)
                    {
                        entity = new TEntity();

                        if (configuration.EntityType.IsValueType)
                        {
                            var obj = (object)entity; // boxing for value types

                            ProcessImport(obj, row, rowIndex, propertyColumnDic, sheetSetting, formulaEvaluator,
                                pictures);

                            entity = (TEntity)obj; // unboxing
                        }
                        else
                        {
                            ProcessImport(entity, row, rowIndex, propertyColumnDic, sheetSetting, formulaEvaluator,
                                pictures);
                        }
                    }
                    else
                    {
                        entity = default;
                    }

                    if (entity is not null)
                    {
                        foreach (var propertyInfo in propertyColumnDic.Keys)
                        {
                            if (propertyInfo.CanWrite)
                            {
                                var propertyValue = propertyInfo.GetValueGetter()?.Invoke(entity);
                                if (InternalCache.InputFormatterFuncCache.TryGetValue(propertyInfo,
                                    out var formatterFunc) && formatterFunc?.Method != null)
                                {
                                    var valueSetter = propertyInfo.GetValueSetter();
                                    if (valueSetter != null)
                                    {
                                        try
                                        {
                                            // apply custom formatterFunc
                                            var formattedValue = formatterFunc.DynamicInvoke(entity, propertyValue);
                                            valueSetter.Invoke(entity, formattedValue);
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
                    }

                    if (configuration.DataFilter?.Invoke(entity) == false)
                    {
                        continue;
                    }

                    dataAction?.Invoke(entity, configuration, rowIndex);

                    yield return entity;
                }
            }
        }
    }

    private static void ProcessImport(object entity, IRow row, int rowIndex,
        Dictionary<PropertyInfo, PropertyConfiguration> propertyColumnDic, SheetSetting sheetSetting,
        IFormulaEvaluator formulaEvaluator,
        Dictionary<CellPosition, IPictureData> pictures)
    {
        foreach (var key in propertyColumnDic.Keys)
        {
            var colIndex = propertyColumnDic[key].ColumnIndex;
            if (colIndex >= 0 && key.CanWrite)
            {
                var columnValue = key.PropertyType.GetDefaultValue();
                var cell = row.GetCell(colIndex);

                if (sheetSetting.CellFilter?.Invoke(cell) != false)
                {
                    var valueSetter = key.GetValueSetter();
                    if (valueSetter != null)
                    {
                        if (key.PropertyType == typeof(byte[])
                            || key.PropertyType == typeof(IPictureData))
                        {
                            if (pictures.TryGetValue(new CellPosition(rowIndex, colIndex), out var pic))
                            {
                                valueSetter.Invoke(entity,
                                    key.PropertyType == typeof(IPictureData) ? pic : pic.Data);
                            }
                        }
                        else
                        {
                            var valueApplied = false;
                            InternalCache.CellReaderFuncCache.TryGetValue(key, out var cellReaderDelegate);
                            string cellValue;
                            if (cellReaderDelegate is Func<ICell, string> cellReader)
                            {
                                cellValue = cellReader.Invoke(cell);
                            }
                            else
                            {
                                cellValue = cell.GetCellValue<string>(formulaEvaluator); 
                            }

                            if (key.PropertyType == typeof(string))
                            {
                                valueApplied = true;
                            }
                            
                            InternalCache.ColumnInputFormatterFuncCache.TryGetValue(key,
                                out var formatterFunc);
                            if (formatterFunc?.Method != null)
                            {
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
                                columnValue = cell.GetCellValue(key.PropertyType, formulaEvaluator);
                            }

                            valueSetter.Invoke(entity, columnValue);
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    ///     Export entity list to excel sheet
    /// </summary>
    /// <typeparam name="TEntity">entity type</typeparam>
    /// <param name="sheet">sheet</param>
    /// <param name="entityList">entity list</param>
    /// <param name="sheetIndex">sheetIndex</param>
    /// <returns>sheet</returns>
    public static ISheet EntitiesToSheet<TEntity>(ISheet sheet, IEnumerable<TEntity>? entityList, int sheetIndex)
    {
        Guard.NotNull(sheet, nameof(sheet));
        if (entityList is null)
        {
            return sheet;
        }

        var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();
        var propertyColumnDictionary = InternalHelper.GetPropertyColumnDictionary(configuration);
        if (propertyColumnDictionary.Keys.Count == 0)
        {
            return sheet;
        }

        var sheetSetting = GetSheetSetting(configuration.SheetSettings, sheetIndex);
        if (sheetSetting.HeaderRowIndex >= 0)
        {
            var headerRow = sheet.CreateRow(sheetSetting.HeaderRowIndex);
            foreach (var key in propertyColumnDictionary.Keys)
            {
                var cell = headerRow.CreateCell(propertyColumnDictionary[key].ColumnIndex);
                cell.SetCellValue(propertyColumnDictionary[key].ColumnTitle);
                sheetSetting.CellAction?.Invoke(cell);
            }

            sheetSetting.RowAction?.Invoke(headerRow);
        }

        var rowIndex = 0;
        foreach (var entity in entityList)
        {
            var row = sheet.CreateRow(sheetSetting.StartRowIndex + rowIndex);
            if (entity != null)
            {
                foreach (var key in propertyColumnDictionary.Keys)
                {
                    var propertyValue = key.GetValueGetter<TEntity>()?.Invoke(entity);
                    if (InternalCache.OutputFormatterFuncCache.TryGetValue(key, out var formatterFunc) &&
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

                    var cell = row.CreateCell(propertyColumnDictionary[key].ColumnIndex);
                    cell.SetCellValue(propertyValue, propertyColumnDictionary[key].ColumnFormatter);
                    sheetSetting.CellAction?.Invoke(cell);
                }
            }

            sheetSetting.RowAction?.Invoke(row);
            rowIndex++;
        }

        PostSheetProcess(sheet, sheetSetting, rowIndex, configuration, propertyColumnDictionary);

        return sheet;
    }

    /// <summary>
    ///     Generic type data table to excel sheet
    /// </summary>
    /// <typeparam name="TEntity">entity type</typeparam>
    /// <param name="sheet">sheet</param>
    /// <param name="dataTable">data table</param>
    /// <param name="sheetIndex">sheetIndex</param>
    /// <returns>sheet</returns>
    public static ISheet DataTableToSheet<TEntity>(ISheet sheet, DataTable? dataTable, int sheetIndex)
    {
        Guard.NotNull(sheet, nameof(sheet));
        if (dataTable is null || dataTable.Rows.Count == 0 || dataTable.Columns.Count == 0)
        {
            return sheet;
        }

        var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();
        var propertyColumnDictionary = InternalHelper.GetPropertyColumnDictionary(configuration);

        if (propertyColumnDictionary.Keys.Count == 0)
        {
            return sheet;
        }

        var sheetSetting = GetSheetSetting(configuration.SheetSettings, sheetIndex);
        if (sheetSetting.HeaderRowIndex >= 0)
        {
            var headerRow = sheet.CreateRow(sheetSetting.HeaderRowIndex);
            for (var i = 0; i < dataTable.Columns.Count; i++)
            {
                var col = propertyColumnDictionary.GetPropertySettingByPropertyName(dataTable.Columns[i]
                    .ColumnName);
                if (null != col)
                {
                    var cell = headerRow.CreateCell(col.ColumnIndex);
                    cell.SetCellValue(col.ColumnTitle);
                    sheetSetting.CellAction?.Invoke(cell);
                }
            }

            sheetSetting.RowAction?.Invoke(headerRow);
        }

        for (var i = 0; i < dataTable.Rows.Count; i++)
        {
            var row = sheet.CreateRow(sheetSetting.StartRowIndex + i);
            for (var j = 0; j < dataTable.Columns.Count; j++)
            {
                var col = propertyColumnDictionary.GetPropertySettingByPropertyName(dataTable.Columns[j]
                    .ColumnName);
                var cell = row.CreateCell(col!.ColumnIndex);
                cell.SetCellValue(dataTable.Rows[i][j], col.ColumnFormatter);
                sheetSetting.CellAction?.Invoke(cell);
            }

            sheetSetting.RowAction?.Invoke(row);
        }

        PostSheetProcess(sheet, sheetSetting, dataTable.Rows.Count, configuration, propertyColumnDictionary);

        return sheet;
    }

    private static void PostSheetProcess<TEntity>(ISheet sheet, SheetSetting sheetSetting, int rowsCount,
        ExcelConfiguration<TEntity> excelConfiguration,
        IDictionary<PropertyInfo, PropertyConfiguration> propertyColumnDictionary)
    {
        if (rowsCount > 0)
        {
            foreach (var setting in propertyColumnDictionary.Values)
            {
                if (setting.ColumnWidth > 0)
                {
                    sheet.SetColumnWidth(setting.ColumnIndex, setting.ColumnWidth * 256);
                }
                else
                {
                    if (sheetSetting.AutoColumnWidthEnabled)
                    {
                        sheet.AutoSizeColumn(setting.ColumnIndex);
                    }
                }
            }

            foreach (var freezeSetting in excelConfiguration.FreezeSettings)
            {
                sheet.CreateFreezePane(freezeSetting.ColSplit, freezeSetting.RowSplit, freezeSetting.LeftMostColumn,
                    freezeSetting.TopRow);
            }

            if (excelConfiguration.FilterSetting != null)
            {
                var headerIndex = sheetSetting.HeaderRowIndex >= 0 ? sheetSetting.HeaderRowIndex : 0;
                sheet.SetAutoFilter(new CellRangeAddress(headerIndex, rowsCount + headerIndex,
                    excelConfiguration.FilterSetting.FirstColumn,
                    excelConfiguration.FilterSetting.LastColumn ??
                    propertyColumnDictionary.Values.Max(_ => _.ColumnIndex)));
            }
        }

        sheetSetting.SheetAction?.Invoke(sheet);
    }
}
