using System;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Attributes
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public sealed class SheetAttribute : Attribute
    {
        public int SheetIndex { get => SheetSetting.SheetIndex; set => SheetSetting.SheetIndex = value; }

        public string SheetName { get => SheetSetting.SheetName; set => SheetSetting.SheetName = value; }

        public int StartRowIndex { get => SheetSetting.StartRowIndex; set => SheetSetting.StartRowIndex = value; }

        public int HeaderRowIndex => SheetSetting.HeaderRowIndex;

        internal SheetSetting SheetSetting { get; }

        public SheetAttribute() => SheetSetting = new SheetSetting();
    }
}
