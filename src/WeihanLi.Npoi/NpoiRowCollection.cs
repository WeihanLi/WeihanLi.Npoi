using System;
using System.Collections;
using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace WeihanLi.Npoi
{
    /// <summary>
    /// Sheet row collection
    /// readonly
    /// </summary>
    public class NpoiRowCollection : IReadOnlyCollection<IRow>
    {
        private readonly ISheet _sheet;

        public NpoiRowCollection(ISheet sheet) => _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));

        public int Count => _sheet.PhysicalNumberOfRows;

        public IEnumerator<IRow> GetEnumerator()
        {
            return (IEnumerator<IRow>)_sheet.GetRowEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
