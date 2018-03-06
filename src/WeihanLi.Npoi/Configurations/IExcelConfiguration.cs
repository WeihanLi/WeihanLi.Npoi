using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    public interface IExcelConfiguration
    {
        IDictionary<PropertyInfo, IPropertyConfiguration> PropertyConfigurationDictionary { get; }

        ExcelSetting ExcelSetting { get; }

        /// <summary>
        /// 设置冻结区域
        /// Creates a split (freezepane). Any existing freezepane or split pane is overwritten.
        /// </summary>
        /// <param name="colSplit">Horizonatal position of split</param>
        /// <param name="rowSplit">Vertical position of split</param>
        IExcelConfiguration HasFreezePane(int colSplit, int rowSplit);

        /// <summary>
        /// 设置冻结区域
        /// Creates a split (freezepane). Any existing freezepane or split pane is overwritten.
        /// </summary>
        /// <param name="colSplit">Horizonatal position of split</param>
        /// <param name="rowSplit">Vertical position of split</param>
        /// <param name="leftmostColumn">Top row visible in bottom pane</param>
        /// <param name="topRow">Left column visible in right pane</param>
        IExcelConfiguration HasFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow);

        /// <summary>
        /// 设置 Filter
        /// </summary>
        /// <param name="firstColumn">firstCol Index of first column</param>
        /// <returns></returns>
        IExcelConfiguration HasFilter(int firstColumn);

        /// <summary>
        /// 设置 Filter
        /// </summary>
        /// <param name="firstColumn">firstCol Index of first column</param>
        /// <param name="lastColumn">lastCol Index of last column (inclusive), must be equal to or larger than {@code firstCol}</param>
        /// <returns></returns>
        IExcelConfiguration HasFilter(int firstColumn, int? lastColumn);
    }

    public interface IExcelConfiguration<TEntity> : IExcelConfiguration
    {
        IPropertyConfiguration Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression);

        #region ExcelSettings FluentAPI

        IExcelConfiguration<TEntity> HasAuthor(string author);

        IExcelConfiguration<TEntity> HasTitle(string title);

        IExcelConfiguration<TEntity> HasDescription(string description);

        IExcelConfiguration<TEntity> HasSubject(string subject);

        #endregion ExcelSettings FluentAPI
    }
}
