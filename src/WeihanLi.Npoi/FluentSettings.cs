// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System.Reflection;
using WeihanLi.Common;
using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;
using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi;

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
        foreach (var type in types.Where(x => x.IsAssignableTo<IMappingProfile>()))
        {
            if (Activator.CreateInstance(type) is IMappingProfile profile)
            {
                LoadMappingProfile(profile);
            }
        }
    }

    /// <summary>
    ///     Load mapping profile for TEntity
    /// </summary>
    /// <typeparam name="TEntity">entity type</typeparam>
    /// <typeparam name="TMappingProfile">entity type mapping profile</typeparam>
    public static void LoadMappingProfile<TEntity, TMappingProfile>()
        where TMappingProfile : IMappingProfile<TEntity>, new()
    {
        var profile = new TMappingProfile();
        profile.Configure(InternalHelper.GetExcelConfigurationMapping<TEntity>());
    }

    /// <summary>
    ///     Load mapping profile for TEntity
    /// </summary>
    /// <param name="profile">profile</param>
    /// <typeparam name="TEntity">entity type</typeparam>
    public static void LoadMappingProfile<TEntity>(IMappingProfile<TEntity> profile)
    {
        Guard.NotNull(profile, nameof(profile));
        profile.Configure(InternalHelper.GetExcelConfigurationMapping<TEntity>());
    }

    /// <summary>
    ///     Load mapping profile for TEntity
    /// </summary>
    /// <typeparam name="TMappingProfile">entity type mapping profile</typeparam>
    public static void LoadMappingProfile<TMappingProfile>() where TMappingProfile : IMappingProfile, new() =>
        LoadMappingProfile(new TMappingProfile());

    /// <summary>
    ///     Load mapping profile for TEntity
    /// </summary>
    /// <param name="profile">profile</param>
    private static void LoadMappingProfile(IMappingProfile profile)
    {
        Guard.NotNull(profile, nameof(profile));
        var profileInterfaceType = profile.GetType()
            .GetImplementedInterfaces()
            .FirstOrDefault(x => x.IsGenericType && x.GetGenericTypeDefinition() == s_profileGenericTypeDefinition);
        if (profileInterfaceType is null)
        {
            return;
        }

        var entityType = profileInterfaceType.GetGenericArguments()[0];
        var configuration = InternalHelper.GetExcelConfigurationMapping(entityType);
        var method = profileInterfaceType.GetMethod(MappingProfileConfigureMethodName,
            new[] { typeof(IExcelConfiguration<>).MakeGenericType(entityType) });
        method?.Invoke(profile, new object[] { configuration });
    }
}
