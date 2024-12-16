// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using WeihanLi.Npoi.Attributes;

namespace WeihanLi.Npoi.Test.Models;

public class OrderTestModel1
{
    public int Id { get; set; }

    public string? Title { get; set; }

    public string? Description { get; set; }
}

public class OrderTestModel2
{
    [Column(Index = 1)]
    public int Id { get; set; }

    [Column(Index = 0)]
    public string? Title { get; set; }

    public string? Description { get; set; }
}

public class OrderTestModel3
{
    public string? Description { get; set; }

    [Column(Index = 0)]
    public int Id { get; set; }

    [Column(Index = 1)]
    public string? Title { get; set; }
}

public class OrderTestModel4
{
    public int Id { get; set; }

    public string? Title { get; set; }

    public string? Description { get; set; }
}
