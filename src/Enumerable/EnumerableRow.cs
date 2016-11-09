using System.Collections;
using System.Collections.Generic;

using NPOI.SS.UserModel;

namespace ExcelFile.net.Enumerable
{
    /// <summary>
    ///     Enumerable rows
    /// </summary>
    public class EnumerableRow : IEnumerable<IRow>
    {
        private readonly ISheet _sheet;
        /// <summary>
        ///     Construct an enumerable rows from a sheet object
        /// </summary>
        /// <param name="sheet">the sheet object</param>
        public EnumerableRow(ISheet sheet)
        {
            _sheet = sheet;
        }
        /// <summary>
        ///     The implementment of IEnumerable
        /// </summary>
        /// <returns></returns>
        public IEnumerator<IRow> GetEnumerator()
        {
            if (_sheet == null
                || _sheet.PhysicalNumberOfRows == 0)
            {
                yield break;
            }
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