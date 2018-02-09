using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using JetBrains.Annotations;
using NPOI.SS.UserModel;
using WeihanLi.Npoi.Attributes;
using WeihanLi.Npoi.Configurations;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi
{
    internal class NpoiHelper<TEntity> where TEntity : new()
    {
        private readonly IDictionary<PropertyInfo, PropertySetting> _propertyColumnDictionary;
        private readonly SheetSetting _sheetSetting;

        internal NpoiHelper()
        {
            _sheetSetting = (typeof(TEntity).GetCustomAttribute<SheetAttribute>() ??
                             new SheetAttribute()).SheetSetting;

            _propertyColumnDictionary = TypeCache.TypePropertySettingDictionary.GetOrAdd(typeof(TEntity),
                GetMapping);
        }

        internal NpoiHelper(ExcelConfiguration<TEntity> excelConfiguration)
        {
            _sheetSetting = (excelConfiguration.SheetSettings[0] as SheetConfiguration)?.SheetSetting;
            _propertyColumnDictionary = excelConfiguration.PropertyConfigurationDictionary.Select(_ =>
                new KeyValuePair<PropertyInfo, PropertySetting>(_.Key,
                    (_.Value as PropertyConfiguration)?.PropertySetting)).ToDictionary(_ => _.Key, _ => _.Value);
        }

        private IDictionary<PropertyInfo, PropertySetting> GetMapping(Type type)
        {
            var dic = new Dictionary<PropertyInfo, PropertySetting>();
            var propertyInfos = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var colIndexList = new List<int>(propertyInfos.Length);
            foreach (var propertyInfo in propertyInfos)
            {
                if (propertyInfo.IsDefined(typeof(IgnoreAttribute)))
                {
                    continue;
                }
                var column = propertyInfo.GetCustomAttribute<ColumnAttribute>() ?? new ColumnAttribute();
                if (column.IsIgnored)
                {
                    continue;
                }
                if (string.IsNullOrWhiteSpace(column.Title))
                {
                    column.Title = propertyInfo.Name;
                }

                // Adjust column index to avoid conflict index
                while (colIndexList.Contains(column.Index))
                {
                    column.Index++;
                }
                colIndexList.Add(column.Index);

                dic.Add(propertyInfo, column.PropertySetting);
            }
            return dic;
        }

        public List<TEntity> SheetToEntityList([NotNull]ISheet sheet)
        {
            var entities = new List<TEntity>();
            var rowEnumerator = sheet.GetRowEnumerator();
            while (rowEnumerator.MoveNext())
            {
                var row = (IRow)rowEnumerator.Current;
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

            // autosize
            for (var i = 0; i < _propertyColumnDictionary.Keys.Count; i++)
            {
                // 自动调整单元格的宽度
                sheet.AutoSizeColumn(i);
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

            // autosize
            for (var i = 0; i < _propertyColumnDictionary.Keys.Count; i++)
            {
                // 自动调整单元格的宽度
                sheet.AutoSizeColumn(i);
            }
            return sheet;
        }
    }
}
