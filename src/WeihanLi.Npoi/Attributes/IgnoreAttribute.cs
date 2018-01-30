using System;

namespace WeihanLi.Npoi.Attributes
{
    /// <summary>
    /// 忽略字段
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class IgnoreAttribute : Attribute
    {
        public bool IsIgnored { get; }

        public IgnoreAttribute(bool isIgnored)
            => IsIgnored = isIgnored;
    }
}
