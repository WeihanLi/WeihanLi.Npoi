using System;

namespace WeihanLi.Npoi.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public sealed class SheetAttribute : Attribute
    {
        public string SheetName { get; }

        public SheetAttribute(string sheetName) => SheetName = sheetName;
    }
}