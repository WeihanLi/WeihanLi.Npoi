// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

namespace WeihanLi.Npoi;

/// <summary>
/// Represents a zero-based row/column coordinate within a sheet.
/// </summary>
/// <param name="Row">Row index.</param>
/// <param name="Column">Column index.</param>
public readonly record struct CellPosition(int Row, int Column);
