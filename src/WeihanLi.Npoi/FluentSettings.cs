using System;
using System.Linq;
using System.Reflection;
using WeihanLi.Common;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi
{
    public static class FluentSettings
    {
        private const string MappingProfileConfigureMethodName = "Configure";
        private static readonly Type s_profileGenericTypeDefinition = typeof(IMappingProfile<>);

        /// <summary>
        ///     Fluent Setting For TEntity
        /// </summary>
        /// <typeparam name="TEntity">TEntity</typeparam>
        /// <returns>excel configuration for entity</returns>
        public static IExcelConfiguration<TEntity> For<TEntity>() =>
            InternalHelper.GetExcelConfigurationMapping<TEntity>();

        /// <summary>
        ///     Load mapping profiles
        /// </summary>
        /// <param name="assemblies">assemblies</param>
        public static void LoadMappingProfiles(params Assembly[] assemblies)
        {
            Guard.NotNull(assemblies, nameof(assemblies));
            if (assemblies.Length == 0)
            {
                assemblies = ReflectHelper.GetAssemblies();
            }

            LoadMappingProfiles(assemblies.SelectMany(ass => ass.GetExportedTypes()).ToArray());
        }

        /// <summary>
        ///     Load mapping profiles
        /// </summary>
        /// <param name="types">mapping profile types</param>
        public static void LoadMappingProfiles(params Type[] types)
        {
            Guard.NotNull(types, nameof(types));
            var profileTypes = types.Where(x => x.IsAssignableTo<IMappingProfile>()).ToArray();
            foreach (var type in profileTypes)
            {
                var profileInterfaceType = type.GetImplementedInterfaces()
                    .FirstOrDefault(
                        x => x.IsGenericType && x.GetGenericTypeDefinition() == s_profileGenericTypeDefinition);
                if (profileInterfaceType is null)
                {
                    continue;
                }

                var profile = Activator.CreateInstance(type);
                var entityType = profileInterfaceType.GetGenericArguments()[0];
                var configuration = InternalHelper.GetExcelConfigurationMapping(entityType);
                var method = profileInterfaceType.GetMethod(MappingProfileConfigureMethodName,
                    new[] {typeof(IExcelConfiguration<>).MakeGenericType(entityType)});
                method?.Invoke(profile, new object[] {configuration});
            }
        }
    }
}
