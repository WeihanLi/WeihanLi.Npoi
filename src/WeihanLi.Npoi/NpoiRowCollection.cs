using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;

namespace WeihanLi.Npoi
{
    /// <summary>
    /// Sheet row collection
    /// </summary>
    public class NpoiRowCollection : IReadOnlyCollection<IRow>
    {
        private readonly ISheet _sheet;

        public NpoiRowCollection(ISheet sheet) => _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));

        public int Count => _sheet.LastRowNum + 1;

        public IEnumerator<IRow> GetEnumerator()
        {
            for (var i = 0; i <= _sheet.LastRowNum; i++)
            {
                yield return _sheet.GetRow(i);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
