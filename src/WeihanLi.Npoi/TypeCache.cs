using System;
using System.Collections.Concurrent;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi
{
    internal class TypeCache
    {
        internal static ConcurrentDictionary<Type, IExcelConfiguration> TypeExcelConfigurationDictionary = new ConcurrentDictionary<Type, IExcelConfiguration>();
    }
}
