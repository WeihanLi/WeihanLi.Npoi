// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System;

namespace WeihanLi.Npoi;

public readonly struct CellPosition : IEquatable<CellPosition>
{
    public CellPosition(int row, int col)
    {
        Row = row;
        Column = col;
    }

    public int Row { get; }
    public int Column { get; }

    public bool Equals(CellPosition other) => Row == other.Row && Column == other.Column;

    public override bool Equals(object? obj) => obj is CellPosition other && Equals(other);

    public override int GetHashCode() => $"{Row}_{Column}".GetHashCode();
}
