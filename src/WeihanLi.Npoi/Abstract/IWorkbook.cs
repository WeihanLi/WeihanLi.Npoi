// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the MIT license.

namespace WeihanLi.Npoi.Abstract;

internal interface IWorkbook
{
    int SheetCount { get; }

    ISheet? GetSheet(int sheetIndex);

    ISheet CreateSheet(string sheetName);

    byte[] ToBytes();
}
