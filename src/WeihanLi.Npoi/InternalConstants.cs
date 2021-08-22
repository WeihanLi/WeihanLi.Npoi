namespace WeihanLi.Npoi
{
    internal static class InternalConstants
    {
        public const int MaxSheetCountXls = 256;
        public const int MaxSheetCountXlsx = 16384;

        public const int MaxRowCountXls = 65536;
        public const int MaxRowCountXlsx = 1_048_576;

        /// <summary>
        ///     DefaultPropertyNameForBasicType
        /// </summary>
        public const string DefaultPropertyNameForBasicType = "Value";

        /// <summary>
        ///     ApplicationName
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
