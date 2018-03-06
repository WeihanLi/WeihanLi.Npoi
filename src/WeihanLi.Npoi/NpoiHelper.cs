using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using JetBrains.Annotations;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using WeihanLi.Npoi.Attributes;
using WeihanLi.Npoi.Configurations;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi
{
    internal class NpoiHelper<TEntity> where TEntity : new()
    {
        private readonly IDictionary<PropertyInfo, PropertySetting> _propertyColumnDictionary;
        private readonly SheetSetting _sheetSetting;
        private readonly ExcelConfiguration<TEntity> _excelConfiguration;

        internal NpoiHelper()
        {
            _excelConfiguration = (ExcelConfiguration<TEntity>)InternalCache.TypeExcelConfigurationDictionary.GetOrAdd(typeof(TEntity),
                GetMapping);

            _sheetSetting = ((SheetConfiguration)_excelConfiguration.SheetConfigurations[0]).SheetSetting;

            //AutoAdjustIndex
            var colIndexList = new List<int>();
            foreach (var item in _excelConfiguration.PropertyConfigurationDictionary.Values.Cast<PropertyConfiguration>().Where(_ => !_.PropertySetting.IsIgnored))
            {
                while (colIndexList.Contains(item.PropertySetting.ColumnIndex))
                {
                    item.PropertySetting.ColumnIndex++;
                }
                colIndexList.Add(item.PropertySetting.ColumnIndex);
            }

            _propertyColumnDictionary = _excelConfiguration.PropertyConfigurationDictionary.Where(_ => !((PropertyConfiguration)_.Value).PropertySetting.IsIgnored).ToDictionary(_ => _.Key, _ => ((PropertyConfiguration)_.Value).PropertySetting);
        }

        private static IExcelConfiguration GetMapping(Type type)
        {
            var excelConfiguration = new ExcelConfiguration<TEntity>
            {
                SheetConfigurations = new ISheetConfiguration[]
                {
                    new SheetConfiguration(type.GetCustomAttribute<SheetAttribute>()?.SheetSetting)
                },
                FilterSetting =
                    type.GetCustomAttribute<FilterAttribute>()?.FilterSeting,
                FreezeSettings =
                    (type.GetCustomAttributes<FreezeAttribute>()?.Select(_ => _.FreezeSetting) ??
                     Enumerable.Empty<FreezeSetting>()).ToList()
            };

            // propertyInfos
            var dic = new Dictionary<PropertyInfo, IPropertyConfiguration>();
            var propertyInfos = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (var propertyInfo in propertyInfos)
            {
                var column = propertyInfo.GetCustomAttribute<ColumnAttribute>() ?? new ColumnAttribute();
                if (string.IsNullOrWhiteSpace(column.Title))
                {
                    column.Title = propertyInfo.Name;
                }
                dic.Add(propertyInfo, new PropertyConfiguration(column.PropertySetting));
            }
            excelConfiguration.PropertyConfigurationDictionary = dic;

            return excelConfiguration;
        }

        public List<TEntity> SheetToEntityList([NotNull]ISheet sheet)
        {
            var entities = new List<TEntity>();
            var rowEnumerator = sheet.GetRowEnumerator();
            while (rowEnumerator.MoveNext())
            {
                if (!(rowEnumerator.Current is IRow row))
                {
                    continue;
                }

                if (row.RowNum == _sheetSetting.HeaderRowIndex) //读取Header
                {
                    for (var i = 0; i < row.Cells.Count; i++)
                    {
                        var col = _propertyColumnDictionary.GetPropertySetting(row.Cells[i].StringCellValue.Trim());
                        if (null != col)
                        {
                            col.ColumnIndex = i;
                        }
                    }
                }
                else if (row.RowNum >= _sheetSetting.StartRowIndex)
                {
                    var entity = new TEntity();
                    for (var i = 0; i < _propertyColumnDictionary.Keys.Count; i++)
                    {
                        var propertyInfo = _propertyColumnDictionary.GetPropertyInfo(i);
                        propertyInfo.SetValue(entity,
                            row.Cells[_propertyColumnDictionary.GetPropertySetting(i).ColumnIndex]
                                .GetCellValue(propertyInfo.PropertyType));
                    }

                    entities.Add(entity);
                }
            }
            return entities;
        }

        public ISheet DataTableToSheet([NotNull]ISheet sheet, DataTable dataTable)
        {
            if (null == dataTable || dataTable.Rows.Count == 0 || dataTable.Columns.Count == 0 || _propertyColumnDictionary.Keys.Count == 0)
            {
                return sheet;
            }
            var headerRow = sheet.CreateRow(0);
            for (var i = 0; i < dataTable.Columns.Count; i++)
            {
                var col = _propertyColumnDictionary.GetPropertySettingByPropertyName(dataTable.Columns[i].ColumnName);
                if (null != col)
                {
                    headerRow.CreateCell(col.ColumnIndex).SetCellValue(col.ColumnTitle);
                }
            }

            for (int i = 0, k = _sheetSetting.StartRowIndex; i < dataTable.Rows.Count; i++)
            {
                var row = sheet.CreateRow(k++);
                for (var j = 0; j < dataTable.Columns.Count; j++)
                {
                    var col = _propertyColumnDictionary.GetPropertySettingByPropertyName(dataTable.Columns[j].ColumnName);
                    row.CreateCell(col.ColumnIndex).SetCellValue(dataTable.Rows[i][j], col.ColumnFormatter);
                }
            }

            // autosizecolumn
            for (var i = 0; i < _propertyColumnDictionary.Values.Max(_ => _.ColumnIndex); i++)
            {
                sheet.AutoSizeColumn(i);
            }

            foreach (var freezeSetting in _excelConfiguration.FreezeSettings)
            {
                sheet.CreateFreezePane(freezeSetting.ColSplit, freezeSetting.RowSplit, freezeSetting.LeftMostColumn, freezeSetting.TopRow);
            }

            if (_excelConfiguration.FilterSetting != null)
            {
                sheet.SetAutoFilter(new CellRangeAddress(_sheetSetting.HeaderRowIndex, dataTable.Rows.Count + _sheetSetting.HeaderRowIndex, _excelConfiguration.FilterSetting.FirstColumn, _excelConfiguration.FilterSetting.LastColumn ?? _propertyColumnDictionary.Values.Max(_ => _.ColumnIndex)));
            }

            return sheet;
        }

        public ISheet EntityListToSheet([NotNull]ISheet sheet, IReadOnlyList<TEntity> entityList)
        {
            if (null == entityList || entityList.Count == 0 || _propertyColumnDictionary.Keys.Count == 0)
            {
                return sheet;
            }

            var headerRow = sheet.CreateRow(_sheetSetting.HeaderRowIndex);
            for (var i = 0; i < _propertyColumnDictionary.Keys.Count; i++)
            {
                var col = _propertyColumnDictionary.GetPropertySetting(i);
                headerRow.CreateCell(col.ColumnIndex).SetCellValue(col.ColumnTitle);
            }

            for (int i = 0, k = _sheetSetting.StartRowIndex; i < entityList.Count; i++)
            {
                var row = sheet.CreateRow(k++);
                for (var j = 0; j < _propertyColumnDictionary.Keys.Count; j++)
                {
                    var property = _propertyColumnDictionary.GetPropertyInfo(j);
                    var col = _propertyColumnDictionary.GetPropertySetting(j);
                    row.CreateCell(col.ColumnIndex).SetCellValue(property.GetValue(entityList[i]), col.ColumnFormatter);
                }
            }

            // AutoSizeColumn
            for (var i = 0; i < _propertyColumnDictionary.Values.Max(_ => _.ColumnIndex); i++)
            {
                sheet.AutoSizeColumn(i);
            }

            foreach (var freezeSetting in _excelConfiguration.FreezeSettings)
            {
                sheet.CreateFreezePane(freezeSetting.ColSplit, freezeSetting.RowSplit, freezeSetting.LeftMostColumn, freezeSetting.TopRow);
            }

            if (_excelConfiguration.FilterSetting != null)
            {
                sheet.SetAutoFilter(new CellRangeAddress(_sheetSetting.HeaderRowIndex, entityList.Count + _sheetSetting.HeaderRowIndex, _excelConfiguration.FilterSetting.FirstColumn, _excelConfiguration.FilterSetting.LastColumn ?? _propertyColumnDictionary.Values.Max(_ => _.ColumnIndex)));
            }

            return sheet;
        }
    }
}
