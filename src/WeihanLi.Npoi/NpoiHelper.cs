using JetBrains.Annotations;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Configurations;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi
{
    internal class NpoiHelper<TEntity> where TEntity : new()
    {
        private readonly IReadOnlyList<SheetSetting> _sheetSettings;
        private readonly IDictionary<PropertyInfo, PropertySetting> _propertyColumnDictionary;
        private readonly ExcelConfiguration<TEntity> _excelConfiguration;
        private readonly Type _entityType;

        internal NpoiHelper()
        {
            _entityType = typeof(TEntity);
            _excelConfiguration = (ExcelConfiguration<TEntity>)InternalCache.TypeExcelConfigurationDictionary.GetOrAdd(_entityType,
                t => InternalHelper.GetExcelConfigurationMapping<TEntity>());

            _sheetSettings = _excelConfiguration.SheetSettings.AsReadOnly();

            InternalHelper.AdjustColumnIndex(_excelConfiguration);

            _propertyColumnDictionary = _excelConfiguration.PropertyConfigurationDictionary.Where(_ => !_.Value.PropertySetting.IsIgnored).ToDictionary(_ => _.Key, _ => _.Value.PropertySetting);
        }

        private SheetSetting GetSheetSetting(int sheetIndex)
        {
            return sheetIndex >= 0 && sheetIndex < _sheetSettings.Count && _sheetSettings.Any(s => s.SheetIndex == sheetIndex)
                ? _sheetSettings.First(s => s.SheetIndex == sheetIndex)
                : _sheetSettings[0];
        }

        public List<TEntity> SheetToEntityList([NotNull]ISheet sheet, int sheetIndex)
        {
            var sheetSetting = GetSheetSetting(sheetIndex);

            var entities = new List<TEntity>(sheet.LastRowNum - sheetSetting.HeaderRowIndex);

            var propertyColumnDic = sheetSetting.HeaderRowIndex >= 0
                ? _propertyColumnDictionary.ToDictionary(_ => _.Key, _ => new PropertySetting()
                {
                    ColumnIndex = -1,
                    ColumnFormatter = _.Value.ColumnFormatter,
                    ColumnTitle = _.Value.ColumnTitle,
                    ColumnWidth = _.Value.ColumnWidth,
                    IsIgnored = _.Value.IsIgnored
                })
                : _propertyColumnDictionary;

            for (var rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (rowIndex == sheetSetting.HeaderRowIndex) // readerHeader
                {
                    if (row != null)
                    {
                        for (var i = 0; i < row.Cells.Count; i++)
                        {
                            if (row.GetCell(i) == null)
                            {
                                continue;
                            }
                            var col = propertyColumnDic.GetPropertySetting(row.GetCell(i).StringCellValue.Trim());
                            if (null != col)
                            {
                                col.ColumnIndex = i;
                            }
                        }
                    }
                    //
                    if (propertyColumnDic.Values.All(_ => _.ColumnIndex < 0))
                    {
                        propertyColumnDic = _propertyColumnDictionary;
                    }
                }
                else if (rowIndex >= sheetSetting.StartRowIndex)
                {
                    if (row == null)
                    {
                        entities.Add(default);
                    }
                    else
                    {
                        TEntity entity;
                        if (row.Cells.Count > 0)
                        {
                            entity = new TEntity();

                            if (typeof(TEntity).IsValueType)
                            {
                                var obj = (object)entity;// boxing for value types
                                foreach (var key in propertyColumnDic.Keys)
                                {
                                    var colIndex = propertyColumnDic[key].ColumnIndex;
                                    if (colIndex >= 0)
                                    {
                                        key.GetValueSetter().Invoke(obj, row.GetCell(colIndex).GetCellValue(key.PropertyType));
                                    }
                                }
                                entity = (TEntity)obj;// unboxing
                            }
                            else
                            {
                                foreach (var key in propertyColumnDic.Keys)
                                {
                                    var colIndex = propertyColumnDic[key].ColumnIndex;
                                    if (colIndex >= 0)
                                    {
                                        key.GetValueSetter().Invoke(entity, row.GetCell(colIndex).GetCellValue(key.PropertyType));
                                    }
                                }
                            }
                        }
                        else
                        {
                            entity = default;
                        }

                        if (null != entity)
                        {
                            foreach (var propertyInfo in _propertyColumnDictionary.Keys)
                            {
                                if (propertyInfo.CanWrite)
                                {
                                    var propertyValue = propertyInfo.GetValueGetter().Invoke(entity);
                                    var formatterFunc = InternalCache.InputFormatterFuncCache.GetOrAdd(propertyInfo, p =>
                                    {
                                        var propertyType = typeof(PropertySetting<,>).MakeGenericType(_entityType, p.PropertyType);
                                        return propertyType.GetProperty("InputFormatterFunc")?.GetValueGetter().Invoke(_propertyColumnDictionary[propertyInfo]);
                                    });
                                    if (null != formatterFunc)
                                    {
                                        var funcType = typeof(Func<,,>).MakeGenericType(_entityType, propertyInfo.PropertyType, typeof(object));
                                        var method = funcType.GetProperty("Method")?.GetValueGetter().Invoke(formatterFunc) as MethodInfo;
                                        var target = funcType.GetProperty("Target")?.GetValueGetter().Invoke(formatterFunc);

                                        if (null != method && target != null)
                                        {
                                            // apply custom formatterFunc
                                            var formattedValue = method.Invoke(target, new[] { entity, propertyValue });
                                            propertyInfo.GetValueSetter().Invoke(entity, formattedValue);
                                        }
                                    }
                                }
                            }
                        }
                        entities.Add(entity);
                    }
                }
            }

            return entities;
        }

        public ISheet DataTableToSheet([NotNull]ISheet sheet, DataTable dataTable, int sheetIndex)
        {
            if (null == dataTable || dataTable.Rows.Count == 0 || dataTable.Columns.Count == 0 || _propertyColumnDictionary.Keys.Count == 0)
            {
                return sheet;
            }
            var sheetSetting = GetSheetSetting(sheetIndex);

            if (sheetSetting.HeaderRowIndex >= 0)
            {
                var headerRow = sheet.CreateRow(sheetSetting.HeaderRowIndex);
                for (var i = 0; i < dataTable.Columns.Count; i++)
                {
                    var col = _propertyColumnDictionary.GetPropertySettingByPropertyName(dataTable.Columns[i].ColumnName);
                    if (null != col)
                    {
                        headerRow.CreateCell(col.ColumnIndex).SetCellValue(col.ColumnTitle);
                    }
                }
            }

            for (var i = 0; i < dataTable.Rows.Count; i++)
            {
                var row = sheet.CreateRow(sheetSetting.StartRowIndex + i);
                for (var j = 0; j < dataTable.Columns.Count; j++)
                {
                    var col = _propertyColumnDictionary.GetPropertySettingByPropertyName(dataTable.Columns[j].ColumnName);
                    row.CreateCell(col.ColumnIndex).SetCellValue(dataTable.Rows[i][j], col.ColumnFormatter);
                }
            }

            PostSheetProcess(sheet, sheetSetting, dataTable.Rows.Count);

            return sheet;
        }

        public ISheet EntityListToSheet([NotNull]ISheet sheet, IReadOnlyList<TEntity> entityList, int sheetIndex)
        {
            if (null == entityList || entityList.Count == 0 || _propertyColumnDictionary.Keys.Count == 0)
            {
                return sheet;
            }
            var sheetSetting = GetSheetSetting(sheetIndex);
            if (sheetSetting.HeaderRowIndex >= 0)
            {
                var headerRow = sheet.CreateRow(sheetSetting.HeaderRowIndex);
                foreach (var key in _propertyColumnDictionary.Keys)
                {
                    headerRow.CreateCell(_propertyColumnDictionary[key].ColumnIndex).SetCellValue(_propertyColumnDictionary[key].ColumnTitle);
                }
            }

            for (var i = 0; i < entityList.Count; i++)
            {
                var row = sheet.CreateRow(sheetSetting.StartRowIndex + i);
                if (null != entityList[i])
                {
                    foreach (var key in _propertyColumnDictionary.Keys)
                    {
                        var propertyValue = key.GetValueGetter<TEntity>().Invoke(entityList[i]);

                        var formatterFunc = InternalCache.OutputFormatterFuncCache.GetOrAdd(key, p =>
                        {
                            var propertyType = typeof(PropertySetting<,>).MakeGenericType(_entityType, p.PropertyType);
                            return propertyType.GetProperty("OutputFormatterFunc")?.GetValueGetter().Invoke(_propertyColumnDictionary[key]);
                        });
                        if (null != formatterFunc)
                        {
                            var funcType = typeof(Func<,,>).MakeGenericType(_entityType, key.PropertyType, typeof(object));
                            var method = funcType.GetProperty("Method")?.GetValueGetter().Invoke(formatterFunc) as MethodInfo;
                            var target = funcType.GetProperty("Target")?.GetValueGetter().Invoke(formatterFunc);

                            if (null != method && target != null)
                            {
                                // apply custom formatterFunc
                                propertyValue = method.Invoke(target, new[] { entityList[i], propertyValue });
                            }
                        }

                        row.CreateCell(_propertyColumnDictionary[key].ColumnIndex).SetCellValue(propertyValue, _propertyColumnDictionary[key].ColumnFormatter);
                    }
                }
            }

            PostSheetProcess(sheet, sheetSetting, entityList.Count);

            return sheet;
        }

        private void PostSheetProcess(ISheet sheet, SheetSetting sheetSetting, int rowsCount)
        {
            // AutoSizeColumn
            foreach (var setting in _propertyColumnDictionary.Values)
            {
                if (setting.ColumnWidth > 0)
                {
                    sheet.SetColumnWidth(setting.ColumnIndex, setting.ColumnWidth * 256);
                }
            }

            foreach (var freezeSetting in _excelConfiguration.FreezeSettings)
            {
                sheet.CreateFreezePane(freezeSetting.ColSplit, freezeSetting.RowSplit, freezeSetting.LeftMostColumn, freezeSetting.TopRow);
            }

            if (_excelConfiguration.FilterSetting != null)
            {
                var headerIndex = sheetSetting.HeaderRowIndex >= 0 ? sheetSetting.HeaderRowIndex : 0;
                sheet.SetAutoFilter(new CellRangeAddress(headerIndex, rowsCount + headerIndex, _excelConfiguration.FilterSetting.FirstColumn, _excelConfiguration.FilterSetting.LastColumn ?? _propertyColumnDictionary.Values.Max(_ => _.ColumnIndex)));
            }
        }
    }
}
