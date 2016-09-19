using System.Collections;
using System.Collections.Generic;

using NPOI.SS.UserModel;

namespace ExcelFile.net.Enumerable
{
    /// <summary>
    ///     Enumerable cells
    /// </summary>
    public class EnumerableCell : IEnumerable<ICell>
    {
        private readonly IRow _row;
        /// <summary>
        ///     Construct an enumerable cells from a row object
        /// </summary>
        /// <param name="row">the row object</param>
        public EnumerableCell(IRow row)
        {
            _row = row;
        }
        /// <summary>
        ///     The implementment of IEnumerable
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ICell> GetEnumerator()
        {
            if (_row == null
                || _row.PhysicalNumberOfCells == 0)
            {
                yield break;
            }
            for (var i = 0; i < _row.LastCellNum; i++)
            {
                yield return _row.GetCell(i);
            }
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}