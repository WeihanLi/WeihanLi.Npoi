using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Reflection;
using WeihanLi.Npoi.Attributes;

namespace WeihanLi.Npoi
{
    public class TypeCache
    {
        public static ConcurrentDictionary<Type, IDictionary<PropertyInfo, ColumnAttribute>> TypeMapCacheDictory =
            new ConcurrentDictionary<Type, IDictionary<PropertyInfo, ColumnAttribute>>();
    }
}
