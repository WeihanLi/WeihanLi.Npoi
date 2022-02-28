// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the MIT license.

using System;

namespace WeihanLi.Npoi.Test.Models;

public class BaseModel
{
    public int Id { get; set; }
}

public class Notice : BaseModel
{
    public string? Title { get; set; }

    public string? Content { get; set; }

    public DateTime PublishedAt { get; set; }

    public string? Publisher { get; set; }
}
