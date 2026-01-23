// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

namespace WeihanLi.Npoi;

internal static class InternalConstants
{
    /// <summary>
    ///     Maximum number of sheets supported by the legacy XLS format.
    /// </summary>
    public const int MaxSheetCountXls = 256;

    /// <summary>
    ///     Maximum number of sheets supported by the XLSX format.
    /// </summary>
    public const int MaxSheetCountXlsx = 16384;

    /// <summary>
    ///     Maximum row count in a single XLS sheet.
    /// </summary>
    public const int MaxRowCountXls = 65536;

    /// <summary>
    ///     Maximum row count in a single XLSX sheet.
    /// </summary>
    public const int MaxRowCountXlsx = 1_048_576;

    /// <summary>
    ///     DefaultPropertyNameForBasicType
    /// </summary>
    public const string DefaultPropertyNameForBasicType = "Value";

    /// <summary>
    ///     ApplicationName
    /// </summary>
    public const string ApplicationName = "WeihanLi.Npoi";

    /// <summary>
    ///     Marker appended when duplicate column titles are encountered.
    /// </summary>
    public const string DuplicateColumnMark = "__dup_mark__";

    #region TemplateParamFormat

    /// <summary>
    ///     Placeholder format for template-wide global parameters.
    /// </summary>
    public const string TemplateGlobalParamFormat = "$(Global:{0})";

    /// <summary>
    ///     Placeholder format for header-level template parameters.
    /// </summary>
    public const string TemplateHeaderParamFormat = "$(Header:{0})";

    /// <summary>
    ///     Placeholder format for body data template parameters.
    /// </summary>
    public const string TemplateDataParamFormat = "$(Data:{0})";

    /// <summary>
    ///     Prefix indicating a template data placeholder.
    /// </summary>
    public const string TemplateDataPrefix = "$(Data:";

    /// <summary>
    ///     Opening tag that surrounds template data sections.
    /// </summary>
    public const string TemplateDataBegin = "<Data>";

    /// <summary>
    ///     Closing tag that surrounds template data sections.
    /// </summary>
    public const string TemplateDataEnd = "</Data>";

    #endregion TemplateParamFormat
}
