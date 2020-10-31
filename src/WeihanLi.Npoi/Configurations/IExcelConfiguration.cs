using System;
using System.Linq.Expressions;

namespace WeihanLi.Npoi.Configurations
{
    public interface IExcelConfiguration
    {
        /// <summary>
        /// Sheet Configuration
        /// </summary>
        /// <param name="sheetIndex">sheetIndex</param>
        /// <param name="sheetName">sheetName</param>
        /// <param name="startRowIndex">startRowIndex</param>
        /// <param name="enableAutoColumnWidth">enable auto column width if true otherwise false</param>
        /// <param name="endRowIndex">endRowIndex, set this if you wanna control where to end(included)</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        [Obsolete("Please use HasSheetConfiguration(int sheetIndex, Action<SheetConfiguration> configAction) override instead")]
        IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex, bool enableAutoColumnWidth, int? endRowIndex = null);

        /// <summary>
        /// Sheet Configuration
        /// </summary>
        /// <param name="configAction">sheet config delegate</param>
        /// <param name="sheetIndex">sheetIndex, 0 is the default value</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        IExcelConfiguration HasSheetConfiguration(Action<SheetConfiguration> configAction, int sheetIndex = 0);

        /// <summary>
        /// setting freeze pane
        /// Creates a split (freeze pane). Any existing freeze pane or split pane is overwritten.
        /// </summary>
        /// <param name="colSplit">Horizontal position of split</param>
        /// <param name="rowSplit">Vertical position of split</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        IExcelConfiguration HasFreezePane(int colSplit, int rowSplit);

        /// <summary>
        /// setting freeze pane
        /// Creates a split (freeze pane). Any existing freeze pane or split pane is overwritten.
        /// </summary>
        /// <param name="colSplit">Horizontal position of split</param>
        /// <param name="rowSplit">Vertical position of split</param>
        /// <param name="leftmostColumn">Top row visible in bottom pane</param>
        /// <param name="topRow">Left column visible in right pane</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        IExcelConfiguration HasFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow);

        /// <summary>
        /// setting filter
        /// </summary>
        /// <param name="firstColumn">firstCol Index of first column</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        IExcelConfiguration HasFilter(int firstColumn);

        /// <summary>
        /// setting filter
        /// </summary>
        /// <param name="firstColumn">firstCol Index of first column</param>
        /// <param name="lastColumn">lastCol Index of last column (inclusive), must be equal to or larger than {@code firstCol}</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        IExcelConfiguration HasFilter(int firstColumn, int? lastColumn);

        #region ExcelSettings FluentAPI

        /// <summary>
        /// set excel Author
        /// </summary>
        /// <param name="author">author</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        IExcelConfiguration HasAuthor(string author);

        /// <summary>
        /// set excel title
        /// </summary>
        /// <param name="title">title</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        IExcelConfiguration HasTitle(string title);

        /// <summary>
        /// set excel description
        /// </summary>
        /// <param name="description">description</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        IExcelConfiguration HasDescription(string description);

        /// <summary>
        /// set excel subject
        /// </summary>
        /// <param name="subject">subject</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        IExcelConfiguration HasSubject(string subject);

        /// <summary>
        /// set excel company
        /// </summary>
        /// <param name="company">company</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        IExcelConfiguration HasCompany(string company);

        /// <summary>
        /// set excel category
        /// </summary>
        /// <param name="category">category</param>
        /// <returns>current excel configuration<see cref="IExcelConfiguration"/></returns>
        IExcelConfiguration HasCategory(string category);

        #endregion ExcelSettings FluentAPI
    }

    public interface IExcelConfiguration<TEntity> : IExcelConfiguration
    {
        /// <summary>
        /// register data validation func
        /// </summary>
        /// <param name="dataValidateFunc">data validate logic</param>
        /// <returns>current excel configuration</returns>
        IExcelConfiguration<TEntity> WithDataValidation(Func<TEntity, bool> dataValidateFunc);

        /// <summary>
        /// property configuration
        /// </summary>
        /// <typeparam name="TProperty">PropertyType</typeparam>
        /// <param name="propertyExpression">propertyExpression to get property info</param>
        /// <returns>current property configuration</returns>
        IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression);

        /// <summary>
        /// property configuration
        /// </summary>
        /// <typeparam name="TProperty">PropertyType</typeparam>
        /// <param name="propertyName">propertyName</param>
        /// <returns>current property configuration</returns>
        IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(string propertyName);
    }
}
