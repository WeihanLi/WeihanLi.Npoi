using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using JetBrains.Annotations;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Attributes;

namespace WeihanLi.Npoi
{
    public static class InternalExtensions
    {
        #region Mapping

        internal static PropertyInfo GetPropertyInfo([NotNull]this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, int index) => mappingDictionary.Keys.ToArray()[index];

        internal static PropertyInfo GetPropertyInfo([NotNull]this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, [NotNull]string propertyName) => mappingDictionary.Keys.FirstOrDefault(k => k.Name.EqualsIgnoreCase(propertyName));

        internal static ColumnAttribute GetColumnAttribute([NotNull]this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, int index) => mappingDictionary.Values.ToArray()[index];

        internal static ColumnAttribute GetColumnAttributeByPropertyName(
            [NotNull] this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, [NotNull] string propertyName)
            => mappingDictionary.Keys.Any(_ => _.Name.EqualsIgnoreCase(propertyName)) ?
                mappingDictionary[mappingDictionary.Keys.First(_ => _.Name.EqualsIgnoreCase(propertyName))] : null;

        internal static ColumnAttribute GetColumnAttribute([NotNull]this IDictionary<PropertyInfo, ColumnAttribute> mappingDictionary, [NotNull]string columnTitle) => mappingDictionary.Values.FirstOrDefault(k => k.Title.EqualsIgnoreCase(columnTitle));

        #endregion Mapping
    }
}
