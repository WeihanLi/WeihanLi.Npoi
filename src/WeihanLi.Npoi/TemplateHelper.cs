// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using WeihanLi.Common;
using WeihanLi.Extensions;

namespace WeihanLi.Npoi;

public sealed class TemplateOptions
{
    /// <summary>
    ///     Global Param Format
    /// </summary>
    public string TemplateGlobalParamFormat
    {
        get;
        set
        {
            if (value.IsNotNullOrWhiteSpace())
            {
                field = value;
            }
        }
    } = InternalConstants.TemplateGlobalParamFormat;

    /// <summary>
    ///     Header Param Format
    /// </summary>
    public string TemplateHeaderParamFormat
    {
        get;
        set
        {
            if (value.IsNotNullOrWhiteSpace())
            {
                field = value;
            }
        }
    } = InternalConstants.TemplateHeaderParamFormat;

    /// <summary>
    ///     Data Param Format
    /// </summary>
    public string TemplateDataParamFormat
    {
        get;
        set
        {
            if (value.IsNotNullOrWhiteSpace())
            {
                field = value;
            }
        }
    } = InternalConstants.TemplateDataParamFormat;

    /// <summary>
    ///     Data Param Prefix
    /// </summary>
    public string TemplateDataPrefix
    {
        get;
        set
        {
            if (value.IsNotNullOrWhiteSpace())
            {
                field = value;
            }
        }
    } = InternalConstants.TemplateDataPrefix;

    /// <summary>
    ///     Data Begin markup
    /// </summary>
    public string TemplateDataBegin
    {
        get;
        set
        {
            if (value.IsNotNullOrWhiteSpace())
            {
                field = value;
            }
        }
    } = InternalConstants.TemplateDataBegin;

    /// <summary>
    ///     Data End markup
    /// </summary>
    public string TemplateDataEnd
    {
        get;
        set
        {
            if (value.IsNotNullOrWhiteSpace())
            {
                field = value;
            }
        }
    } = InternalConstants.TemplateDataEnd;
}

public static class TemplateHelper
{
    /// <summary>
    ///     Configure TemplateOptions
    /// </summary>
    /// <param name="optionsAction">optionsAction</param>
    public static void ConfigureTemplateOptions(Action<TemplateOptions> optionsAction)
    {
        Guard.NotNull(optionsAction);

        optionsAction.Invoke(NpoiTemplateHelper.s_templateOptions);
    }
}
