using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi
{
    public class FluentSettings
    {
        /// <summary>
        /// Fluent Setting For TEntity
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <returns></returns>
        public static IExcelConfiguration<TEntity> For<TEntity>() => InternalHelper.GetExcelConfigurationMapping<TEntity>();
    }
}
