namespace WeihanLi.Npoi
{
    internal static class InternalConstants
    {
        /// <summary>
        /// Workbook中最多Sheet
        /// </summary>
        public const int MaxSheetNum = 256;

        /// <summary>
        /// DefaultPropertyNameForBasicType
        /// </summary>
        public const string DefaultPropertyNameForBasicType = "Value";

        /// <summary>
        /// ApplicationName
        /// </summary>
        public const string ApplicationName = "WeihanLi.Npoi";

        #region TemplateParamFormat

        public const string TemplateGlobalParamFormat = "$(Global:{0})";
        public const string TemplateHeaderParamFormat = "$(Header:{0})";
        public const string TemplateDataParamFormat = "$(Data:{0})";
        public const string TemplateDataPrefix = "$(Data:";
        public const string TemplateDataBegin = "<Data>";
        public const string TemplateDataEnd = "</Data>";

        #endregion TemplateParamFormat
    }
}
