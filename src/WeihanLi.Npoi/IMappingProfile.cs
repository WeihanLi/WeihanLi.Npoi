// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using WeihanLi.Npoi.Configurations;

namespace WeihanLi.Npoi;

/// <summary>
/// Marker interface for describing fluent configuration profiles.
/// </summary>
public interface IMappingProfile;

/// <summary>
/// Strongly typed mapping profile contract.
/// </summary>
/// <typeparam name="T">Entity type being configured.</typeparam>
public interface IMappingProfile<T> : IMappingProfile
{
    /// <summary>
    ///     Configures the Excel mapping metadata for the given entity type.
    /// </summary>
    /// <param name="configuration">Excel configuration builder.</param>
    void Configure(IExcelConfiguration<T> configuration);
}
