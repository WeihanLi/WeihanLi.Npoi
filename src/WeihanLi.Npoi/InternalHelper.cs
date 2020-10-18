using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using WeihanLi.Npoi.Attributes;
using WeihanLi.Npoi.Configurations;

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
                  FilterSetting = type.GetCustomAttribute<FilterAttribute>()?.FilterSetting
              };
              foreach (var sheetAttribute in type.GetCustomAttributes<SheetAttribute>())
              {
                  if (sheetAttribute.SheetIndex >= 0)
                  {
                      excelConfiguration.SheetSettings[sheetAttribute.SheetIndex] = sheetAttribute.SheetSetting;
                  }
              }
              foreach (var freezeAttribute in type.GetCustomAttributes<FreezeAttribute>())
              {
                  excelConfiguration.FreezeSettings.Add(freezeAttribute.FreezeSetting);
              }
              var dic = new Dictionary<PropertyInfo, PropertyConfiguration>();
              var propertyInfos = Common.CacheUtil.GetTypeProperties(type);
              foreach (var propertyInfo in propertyInfos)
              {
                  var column = propertyInfo.GetCustomAttribute<ColumnAttribute>() ?? new ColumnAttribute();
                  if (string.IsNullOrWhiteSpace(column.Title))
                  {
                      column.Title = propertyInfo.Name;
                  }

                  var propertyConfigurationType =
                      typeof(PropertyConfiguration<,>).MakeGenericType(type, propertyInfo.PropertyType);
                  var propertyConfiguration = Activator.CreateInstance(propertyConfigurationType, new object[] { propertyInfo });

                  propertyConfigurationType.GetProperty(nameof(column.PropertyConfiguration.ColumnTitle))?.GetSetMethod()?
                      .Invoke(propertyConfiguration, new object[] { column.PropertyConfiguration.ColumnTitle });
                  propertyConfigurationType.GetProperty(nameof(column.PropertyConfiguration.ColumnIndex))?.GetSetMethod()?
                      .Invoke(propertyConfiguration, new object[] { column.PropertyConfiguration.ColumnIndex });
                  propertyConfigurationType.GetProperty(nameof(column.PropertyConfiguration.ColumnFormatter))?.GetSetMethod()?
                      .Invoke(propertyConfiguration, new object[] { column.PropertyConfiguration.ColumnFormatter });
                  propertyConfigurationType.GetProperty(nameof(column.PropertyConfiguration.IsIgnored))?.GetSetMethod()?
                      .Invoke(propertyConfiguration, new object[] { column.PropertyConfiguration.IsIgnored });
                  propertyConfigurationType.GetProperty(nameof(column.PropertyConfiguration.ColumnWidth))?.GetSetMethod()?
                      .Invoke(propertyConfiguration, new object[] { column.PropertyConfiguration.ColumnWidth });

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
            ICollection<int> validColumnIndex = excelConfiguration.PropertyConfigurationDictionary.Values.
                Where(c => !c.IsIgnored)
                .Select(x => x.ColumnIndex)
                .ToArray();
            if (validColumnIndex.All(_ => _ >= 0) &&
                validColumnIndex.Distinct().Count() == validColumnIndex.Count)
            {
                return;
            }

            var colIndexList = new List<int>(validColumnIndex.Count);
            foreach (var item in excelConfiguration.PropertyConfigurationDictionary
                .Where(_ => !_.Value.IsIgnored)
                .OrderBy(_ => _.Value.ColumnIndex >= 0 ? _.Value.ColumnIndex : int.MaxValue)
                .ThenBy(_ => _.Key.Name)
                .Select(_ => _.Value)
                )
            {
                while (colIndexList.Contains(item.ColumnIndex) || item.ColumnIndex < 0)
                {
                    if (colIndexList.Count > 0)
                    {
                        item.ColumnIndex = colIndexList.Max() + 1;
                    }
                    else
                    {
                        item.ColumnIndex++;
                    }
                }
                colIndexList.Add(item.ColumnIndex);
            }
        }

        /// <summary>
        /// GetPropertyColumnDictionary
        /// </summary>
        /// <typeparam name="TEntity">TEntity Type</typeparam>
        /// <returns></returns>
        public static Dictionary<PropertyInfo, PropertyConfiguration> GetPropertyColumnDictionary<TEntity>() => GetPropertyColumnDictionary(GetExcelConfigurationMapping<TEntity>());

        /// <summary>
        /// GetPropertyColumnDictionary
        /// </summary>
        /// <typeparam name="TEntity">TEntity Type</typeparam>
        /// <returns></returns>
        public static Dictionary<PropertyInfo, PropertyConfiguration> GetPropertyColumnDictionary<TEntity>(ExcelConfiguration<TEntity> configuration)
        {
            AdjustColumnIndex(configuration);

            return configuration.PropertyConfigurationDictionary
                .Where(p => !p.Value.IsIgnored)
                .ToDictionary(_ => _.Key, _ => _.Value);
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
                .Where(p => !p.Value.IsIgnored)
                .OrderBy(p => p.Value.ColumnIndex)
                .Select(p => p.Key)
                .ToArray();
        }
    }
}
