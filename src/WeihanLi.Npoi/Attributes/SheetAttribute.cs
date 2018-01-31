using System;

namespace WeihanLi.Npoi.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public sealed class SheetAttribute : Attribute
    {
        public string SheetName { get; }

        public int StartRowIndex { get; set; } = 1;

        public int HeaderRowIndex => StartRowIndex - 1;

        public SheetAttribute(string sheetName) => SheetName = sheetName;
    }
}
