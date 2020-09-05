using System;
using WeihanLi.Document.Settings.Excel;

namespace WeihanLi.Document.Attributes.Excel
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public sealed class SheetAttribute : Attribute
    {
        public int SheetIndex { get => SheetSetting.SheetIndex; set => SheetSetting.SheetIndex = value; }

        public string SheetName { get => SheetSetting.SheetName; set => SheetSetting.SheetName = value; }

        public int StartRowIndex { get => SheetSetting.StartRowIndex; set => SheetSetting.StartRowIndex = value; }

        public int HeaderRowIndex => SheetSetting.HeaderRowIndex;

        public bool AutoColumnWidthEnabled
        {
            get => SheetSetting.AutoColumnWidthEnabled;
            set => SheetSetting.AutoColumnWidthEnabled = value;
        }

        internal SheetSetting SheetSetting { get; }

        public SheetAttribute() => SheetSetting = new SheetSetting();
    }
}
