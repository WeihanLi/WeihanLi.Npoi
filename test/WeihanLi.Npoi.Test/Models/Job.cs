// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System.ComponentModel.DataAnnotations;

namespace WeihanLi.Npoi.Test.Models;

public sealed record Job
{
    public int Id { get; set; }

    [Required]
    public string Name { get; set; } = default!;
}
