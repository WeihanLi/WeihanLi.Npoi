using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Reflection;
using WeihanLi.Npoi.Attributes;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi
{
    internal class TypeCache
    {
        internal static ConcurrentDictionary<Type, IDictionary<PropertyInfo, ColumnAttribute>> TypeMapCacheDictionary =
            new ConcurrentDictionary<Type, IDictionary<PropertyInfo, ColumnAttribute>>();

        internal static ConcurrentDictionary<Type, IDictionary<PropertyInfo, PropertySetting>> TypePropertySettingDictionary = new ConcurrentDictionary<Type, IDictionary<PropertyInfo, PropertySetting>>();
    }
}
