﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using JetBrains.Annotations;
using NPOI.SS.UserModel;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Attributes;

namespace WeihanLi.Npoi
{
    internal class NpoiHelper<TEntity> where TEntity : new()
    {
        private IDictionary<PropertyInfo, ColumnAttribute> _propertyColumnDictionary;

        internal NpoiHelper()
        {
            _propertyColumnDictionary = TypeCache.TypeMapCacheDictory.GetOrAdd(typeof(TEntity),
                GetMapping);
        }

        private IDictionary<PropertyInfo, ColumnAttribute> GetMapping(Type type)
        {
            var dic = new Dictionary<PropertyInfo, ColumnAttribute>();
            var propertyInfos = type.GetProperties(BindingFlags.Public | BindingFlags.Instance).ToList();
            // TODO: Adjust column index
            foreach (var propertyInfo in propertyInfos)
            {
                var attribute = propertyInfo.GetCustomAttribute<ColumnAttribute>();
                if (null != attribute)
                {
                    dic.Add(propertyInfo, attribute);
                }
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
                if (row.RowNum == 0) //读取Header
                {
                    for (var i = 0; i < row.Cells.Count; i++)
                    {
                        var col = _propertyColumnDictionary.GetColumnAttribute(row.Cells[i].StringCellValue);
                        if (null != col)
                        {
                            col.Index = i;
                        }
                    }
                }
                else
                {
                    var entity = new TEntity();
                    for (var i = 0; i < _propertyColumnDictionary.Keys.Count; i++)
                    {
                        _propertyColumnDictionary.GetPropertyInfo(i).SetValue(entity,
                            row.Cells[_propertyColumnDictionary.GetColumnAttribute(i).Index]
                                .GetCellValue(_propertyColumnDictionary.GetPropertyInfo(i).PropertyType));
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
                var col = _propertyColumnDictionary.GetColumnAttributeByPropertyName(dataTable.Columns[i].ColumnName);
                if (null != col)
                {
                    headerRow.CreateCell(col.Index).SetCellValue(col.Title);
                }
            }

            for (var i = 0; i < dataTable.Rows.Count; i++)
            {
                var row = sheet.CreateRow(i + 1);
                for (var j = 0; j < dataTable.Columns.Count; j++)
                {
                    row.CreateCell(_propertyColumnDictionary.GetColumnAttributeByPropertyName(dataTable.Columns[j].ColumnName).Index).SetCellValue(dataTable.Rows[i][j]);
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

            var headerRow = sheet.CreateRow(0);

            for (var i = 0; i < _propertyColumnDictionary.Keys.Count; i++)
            {
                headerRow.CreateCell(_propertyColumnDictionary.GetColumnAttribute(i).Index).SetCellValue(_propertyColumnDictionary.GetColumnAttribute(i).Title);
            }

            for (var i = 0; i < entityList.Count; i++)
            {
                var row = sheet.CreateRow(i + 1);
                for (var j = 0; j < _propertyColumnDictionary.Keys.Count; j++)
                {
                    var property = _propertyColumnDictionary.GetPropertyInfo(j);
                    row.CreateCell(_propertyColumnDictionary.GetColumnAttribute(j).Index).SetCellValue(property.GetValue(entityList[i]));
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
