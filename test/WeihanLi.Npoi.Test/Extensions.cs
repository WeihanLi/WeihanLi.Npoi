// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

using System.Data;
using Xunit;

namespace WeihanLi.Npoi.Test;

public static class Extensions
{
    public static DataRow AddNewRow(this DataTable datatable, object[]? rowData = null)
    {
        var row = datatable.NewRow();
        if (rowData != null)
        {
            row.ItemArray = rowData;
        }
        datatable.Rows.Add(row);
        return row;
    }

    public static void AssertEquals(this DataTable actual, DataTable expected)
    {
        // Check columns
        for (var headerIndex = 0; headerIndex < expected.Columns.Count; headerIndex++)
        {
            var expectedValue = expected.Columns[headerIndex].ToString();
            var excelValue = actual.Columns[headerIndex].ToString();

            // "TRUE" from header column is translated to "True".
            // I don't know how to load display value of boolean, therefore I ignore letter casing.
            Assert.Equal(expectedValue, excelValue, ignoreCase: true);
        }

        // Check rows
        for (var rowIndex = 0; rowIndex < expected.Rows.Count; rowIndex++)
        {
            for (var colIndex = 0; colIndex < expected.Rows[rowIndex].ItemArray.Length; colIndex++)
            {
                var expectedValue = expected.Rows[rowIndex].ItemArray[colIndex]?.ToString();
                var excelValue = actual.Rows[rowIndex][colIndex].ToString();
                Assert.Equal(expectedValue, excelValue);
            }
        }
    }

}
