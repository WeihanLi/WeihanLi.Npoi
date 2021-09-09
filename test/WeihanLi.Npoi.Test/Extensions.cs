using System.Data;

namespace WeihanLi.Npoi.Test
{
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
    }
}
