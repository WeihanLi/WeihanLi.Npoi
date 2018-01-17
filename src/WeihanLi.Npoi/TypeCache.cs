using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Reflection;
using WeihanLi.Npoi.Attributes;

namespace WeihanLi.Npoi
{
    internal class TypeCache
    {
        internal static ConcurrentDictionary<Type, IDictionary<PropertyInfo, ColumnAttribute>> TypeMapCacheDictory =
            new ConcurrentDictionary<Type, IDictionary<PropertyInfo, ColumnAttribute>>();
    }
}
