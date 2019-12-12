using System;
using System.Linq.Expressions;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    public interface IExcelConfiguration
    {
        ExcelSetting ExcelSetting { get; }

        IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName);

        IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, bool enableAutoColumnWidth);

        IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex);

        IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex, bool enableAutoColumnWidth);

        /// <summary>
        /// setting freeze pane
        /// Creates a split (freeze pane). Any existing freeze pane or split pane is overwritten.
        /// </summary>
        /// <param name="colSplit">Horizontal position of split</param>
        /// <param name="rowSplit">Vertical position of split</param>
        IExcelConfiguration HasFreezePane(int colSplit, int rowSplit);

        /// <summary>
        /// setting freeze pane
        /// Creates a split (freeze pane). Any existing freeze pane or split pane is overwritten.
        /// </summary>
        /// <param name="colSplit">Horizontal position of split</param>
        /// <param name="rowSplit">Vertical position of split</param>
        /// <param name="leftmostColumn">Top row visible in bottom pane</param>
        /// <param name="topRow">Left column visible in right pane</param>
        IExcelConfiguration HasFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow);

        /// <summary>
        /// setting filter
        /// </summary>
        /// <param name="firstColumn">firstCol Index of first column</param>
        /// <returns></returns>
        IExcelConfiguration HasFilter(int firstColumn);

        /// <summary>
        /// setting filter
        /// </summary>
        /// <param name="firstColumn">firstCol Index of first column</param>
        /// <param name="lastColumn">lastCol Index of last column (inclusive), must be equal to or larger than {@code firstCol}</param>
        /// <returns></returns>
        IExcelConfiguration HasFilter(int firstColumn, int? lastColumn);

        #region ExcelSettings FluentAPI

        /// <summary>
        /// set excel Author
        /// </summary>
        /// <param name="author">author</param>
        /// <returns></returns>
        IExcelConfiguration HasAuthor(string author);

        /// <summary>
        /// set excel title
        /// </summary>
        /// <param name="title">title</param>
        /// <returns></returns>
        IExcelConfiguration HasTitle(string title);

        /// <summary>
        /// set excel description
        /// </summary>
        /// <param name="description">description</param>
        /// <returns></returns>
        IExcelConfiguration HasDescription(string description);

        /// <summary>
        /// set excel subject
        /// </summary>
        /// <param name="subject">subject</param>
        /// <returns></returns>
        IExcelConfiguration HasSubject(string subject);

        /// <summary>
        /// set excel company
        /// </summary>
        /// <param name="company">company</param>
        /// <returns></returns>
        IExcelConfiguration HasCompany(string company);

        /// <summary>
        /// set excel category
        /// </summary>
        /// <param name="category">category</param>
        /// <returns></returns>
        IExcelConfiguration HasCategory(string category);

        #endregion ExcelSettings FluentAPI
    }

    public interface IExcelConfiguration<TEntity> : IExcelConfiguration
    {
        IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression);

        IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(string propertyName);
    }
}
