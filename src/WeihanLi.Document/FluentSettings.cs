using WeihanLi.Document.Configurations;
using WeihanLi.Document.Configurations.Excel;

namespace WeihanLi.Document
{
    public class FluentSettings
    {
        /// <summary>
        /// Fluent Setting For TEntity
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <returns></returns>
        public static IExcelConfiguration<TEntity> ExcelSettingsFor<TEntity>() => InternalHelper.GetExcelConfigurationMapping<TEntity>();

        /// <summary>
        /// Fluent Setting For TEntity
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <returns></returns>
        public static IDocumentConfiguration<TEntity> For<TEntity>() => InternalHelper.GetExcelConfigurationMapping<TEntity>();
    }
}
