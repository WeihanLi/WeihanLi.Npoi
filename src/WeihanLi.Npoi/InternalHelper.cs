using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using WeihanLi.Npoi.Attributes;
using WeihanLi.Npoi.Configurations;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi
{
    internal static class InternalHelper
    {
        /// <summary>
        /// GetExcelConfigurationMapping
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <returns>IExcelConfiguration</returns>
        public static ExcelConfiguration<TEntity> GetExcelConfigurationMapping<TEntity>()
        {
            var configuration = (ExcelConfiguration<TEntity>)InternalCache.TypeExcelConfigurationDictionary.GetOrAdd(typeof(TEntity), type =>
          {
              var excelConfiguration = new ExcelConfiguration<TEntity>
              {
                  SheetSettings = new Dictionary<int, SheetSetting>()
              {
                    { 0, new SheetSetting() },
              },
                  FilterSetting = type.GetCustomAttribute<FilterAttribute>()?.FilterSetting,
                  FreezeSettings = type.GetCustomAttributes<FreezeAttribute>().Select(_ => _.FreezeSetting).ToList()
              };
              foreach (var sheetAttribute in type.GetCustomAttributes<SheetAttribute>())
              {
                  if (sheetAttribute.SheetIndex > 0)
                  {
                      excelConfiguration.SheetSettings[sheetAttribute.SheetIndex] = sheetAttribute.SheetSetting;
                  }
              }
              var dic = new Dictionary<PropertyInfo, PropertyConfiguration>();
              var propertyInfos = Common.CacheUtil.TypePropertyCache.GetOrAdd(type, t => t.GetProperties());
              foreach (var propertyInfo in propertyInfos)
              {
                  var column = propertyInfo.GetCustomAttribute<ColumnAttribute>() ?? new ColumnAttribute();
                  if (string.IsNullOrWhiteSpace(column.Title))
                  {
                      column.Title = propertyInfo.Name;
                  }

                  var propertySettingType = typeof(PropertySetting<,>).MakeGenericType(type, propertyInfo.PropertyType);
                  var propertySetting = Activator.CreateInstance(propertySettingType);

                  propertySettingType.GetProperty(nameof(column.PropertySetting.ColumnTitle))?.GetSetMethod()?
                      .Invoke(propertySetting, new object[] { column.PropertySetting.ColumnTitle });
                  propertySettingType.GetProperty(nameof(column.PropertySetting.ColumnIndex))?.GetSetMethod()?
                      .Invoke(propertySetting, new object[] { column.PropertySetting.ColumnIndex });
                  propertySettingType.GetProperty(nameof(column.PropertySetting.ColumnFormatter))?.GetSetMethod()?
                      .Invoke(propertySetting, new object[] { column.PropertySetting.ColumnFormatter });
                  propertySettingType.GetProperty(nameof(column.PropertySetting.IsIgnored))?.GetSetMethod()?
                      .Invoke(propertySetting, new object[] { column.PropertySetting.IsIgnored });
                  propertySettingType.GetProperty(nameof(column.PropertySetting.ColumnWidth))?.GetSetMethod()?
                      .Invoke(propertySetting, new object[] { column.PropertySetting.ColumnWidth });

                  var propertyConfigurationType =
                      typeof(PropertyConfiguration<,>).MakeGenericType(type, propertyInfo.PropertyType);
                  var propertyConfiguration = Activator.CreateInstance(propertyConfigurationType, new object[] { propertySetting });

                  dic.Add(propertyInfo, (PropertyConfiguration)propertyConfiguration);
              }
              excelConfiguration.PropertyConfigurationDictionary = dic;

              return excelConfiguration;
          });
            return configuration;
        }

        /// <summary>
        /// Adjust Column Index
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <param name="excelConfiguration">excelConfiguration</param>
        private static void AdjustColumnIndex<TEntity>(ExcelConfiguration<TEntity> excelConfiguration)
        {
            if (excelConfiguration.PropertyConfigurationDictionary.Values.All(_ => _.PropertySetting.ColumnIndex >= 0) &&
                excelConfiguration.PropertyConfigurationDictionary.Values.Select(_ => _.PropertySetting.ColumnIndex)
                    .Distinct().Count() == excelConfiguration.PropertyConfigurationDictionary.Values.Count)
            {
                return;
            }

            var colIndexList = new List<int>(excelConfiguration.PropertyConfigurationDictionary.Count);
            foreach (var item in excelConfiguration.PropertyConfigurationDictionary
                .Where(_ => !_.Value.PropertySetting.IsIgnored)
                .OrderBy(_ => _.Value.PropertySetting.ColumnIndex >= 0 ? _.Value.PropertySetting.ColumnIndex : int.MaxValue)
                .ThenBy(_ => _.Key.Name)
                .Select(_ => _.Value)
                )
            {
                while (colIndexList.Contains(item.PropertySetting.ColumnIndex) || item.PropertySetting.ColumnIndex < 0)
                {
                    if (colIndexList.Count > 0)
                    {
                        item.PropertySetting.ColumnIndex = colIndexList.Max() + 1;
                    }
                    else
                    {
                        item.PropertySetting.ColumnIndex++;
                    }
                }
                colIndexList.Add(item.PropertySetting.ColumnIndex);
            }
        }

        /// <summary>
        /// GetPropertyColumnDictionary
        /// </summary>
        /// <typeparam name="TEntity">TEntity Type</typeparam>
        /// <returns></returns>
        public static Dictionary<PropertyInfo, PropertySetting> GetPropertyColumnDictionary<TEntity>() => GetPropertyColumnDictionary(GetExcelConfigurationMapping<TEntity>());

        /// <summary>
        /// GetPropertyColumnDictionary
        /// </summary>
        /// <typeparam name="TEntity">TEntity Type</typeparam>
        /// <returns></returns>
        public static Dictionary<PropertyInfo, PropertySetting> GetPropertyColumnDictionary<TEntity>(ExcelConfiguration<TEntity> configuration)
        {
            AdjustColumnIndex(configuration);

            return configuration.PropertyConfigurationDictionary
                .Where(p => !p.Value.PropertySetting.IsIgnored)
                .ToDictionary(_ => _.Key, _ => _.Value.PropertySetting);
        }

        /// <summary>
        /// GetProperties
        /// </summary>
        /// <typeparam name="TEntity">TEntity Type</typeparam>
        /// <returns></returns>
        public static IReadOnlyList<PropertyInfo> GetPropertiesForCsvHelper<TEntity>()
        {
            var configuration = GetExcelConfigurationMapping<TEntity>();

            AdjustColumnIndex(configuration);

            return configuration.PropertyConfigurationDictionary
                .Where(p => !p.Value.PropertySetting.IsIgnored)
                .OrderBy(p => p.Value.PropertySetting.ColumnIndex)
                .ThenBy(p => p.Key.Name)
                .Select(p => p.Key)
                .ToArray();
        }
    }
}
