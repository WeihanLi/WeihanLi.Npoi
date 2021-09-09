using System;
using System.Linq.Expressions;
using WeihanLi.Npoi.Settings;

namespace WeihanLi.Npoi.Configurations
{
    public interface IExcelConfiguration
    {
        /// <summary>
        ///     Sheet Configuration
        /// </summary>
        /// <param name="configAction">sheet config delegate</param>
        /// <param name="sheetIndex">sheetIndex, 0 is the default value</param>
        /// <returns>current excel configuration<see ref="IExcelConfiguration" /></returns>
        IExcelConfiguration HasSheetSetting(Action<SheetSetting> configAction, int sheetIndex = 0);

        /// <summary>
        ///     excel setting configure
        /// </summary>
        /// <param name="configAction">config delegate</param>
        /// <returns>current excel configuration<see ref="IExcelConfiguration" /></returns>
        IExcelConfiguration HasExcelSetting(Action<ExcelSetting> configAction);

        /// <summary>
        ///     setting freeze pane
        ///     Creates a split (freeze pane). Any existing freeze pane or split pane is overwritten.
        /// </summary>
        /// <param name="colSplit">Horizontal position of split</param>
        /// <param name="rowSplit">Vertical position of split</param>
        /// <returns>current excel configuration<see ref="IExcelConfiguration" /></returns>
        IExcelConfiguration HasFreezePane(int colSplit, int rowSplit);

        /// <summary>
        ///     setting freeze pane
        ///     Creates a split (freeze pane). Any existing freeze pane or split pane is overwritten.
        /// </summary>
        /// <param name="colSplit">Horizontal position of split</param>
        /// <param name="rowSplit">Vertical position of split</param>
        /// <param name="leftmostColumn">Top row visible in bottom pane</param>
        /// <param name="topRow">Left column visible in right pane</param>
        /// <returns>current excel configuration<see ref="IExcelConfiguration" /></returns>
        IExcelConfiguration HasFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow);

        /// <summary>
        ///     setting filter
        /// </summary>
        /// <param name="firstColumn">firstCol Index of first column</param>
        /// <returns>current excel configuration<see ref="IExcelConfiguration" /></returns>
        IExcelConfiguration HasFilter(int firstColumn);

        /// <summary>
        ///     setting filter
        /// </summary>
        /// <param name="firstColumn">firstCol Index of first column</param>
        /// <param name="lastColumn">lastCol Index of last column (inclusive), must be equal to or larger than {@code firstCol}</param>
        /// <returns>current excel configuration<see ref="IExcelConfiguration" /></returns>
        IExcelConfiguration HasFilter(int firstColumn, int? lastColumn);
    }

    public interface IExcelConfiguration<TEntity> : IExcelConfiguration
    {
        /// <summary>
        ///     register data validation func
        /// </summary>
        /// <param name="dataValidateFunc">data validate logic</param>
        /// <returns>current excel configuration</returns>
        IExcelConfiguration<TEntity> WithDataValidation(Func<TEntity?, bool>? dataValidateFunc);

        /// <summary>
        ///     property configuration
        /// </summary>
        /// <typeparam name="TProperty">PropertyType</typeparam>
        /// <param name="propertyExpression">propertyExpression to get property info</param>
        /// <returns>current property configuration</returns>
        IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(
            Expression<Func<TEntity, TProperty>> propertyExpression);

        /// <summary>
        ///     property configuration
        /// </summary>
        /// <typeparam name="TProperty">PropertyType</typeparam>
        /// <param name="propertyName">propertyName</param>
        /// <returns>current property configuration</returns>
        IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(string propertyName);
    }
}
