using JetBrains.Annotations;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi
{
    internal static class InternalExtensions
    {
        /// <summary>
        /// GetPropertySettingByPropertyName
        /// </summary>
        /// <param name="mappingDictionary">mappingDictionary</param>
        /// <param name="propertyName">propertyName</param>
        /// <returns></returns>
        internal static PropertySetting GetPropertySettingByPropertyName([NotNull] this IDictionary<PropertyInfo, PropertySetting> mappingDictionary, [NotNull] string propertyName)
            => mappingDictionary.Keys.Any(_ => _.Name.EqualsIgnoreCase(propertyName)) ?
                mappingDictionary[mappingDictionary.Keys.First(_ => _.Name.EqualsIgnoreCase(propertyName))] : null;

        /// <summary>
        /// GetPropertySettingByColumnName
        /// </summary>
        /// <param name="mappingDictionary">mappingDictionary</param>
        /// <param name="columnTitle">columnTitle</param>
        /// <returns></returns>
        internal static PropertySetting GetPropertySetting([NotNull]this IDictionary<PropertyInfo, PropertySetting> mappingDictionary, [NotNull]string columnTitle) => mappingDictionary.Values.FirstOrDefault(k => k.ColumnTitle.EqualsIgnoreCase(columnTitle));
    }
}
