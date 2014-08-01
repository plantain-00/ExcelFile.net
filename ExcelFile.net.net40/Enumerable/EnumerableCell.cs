using System.Collections;
using System.Collections.Generic;

using NPOI.SS.UserModel;

namespace ExcelFile.net.Enumerable
{
    /// <summary>
    ///     某一行的单元格枚举器
    /// </summary>
    public class EnumerableCell : IEnumerable<ICell>
    {
        private readonly IRow _row;
        /// <summary>
        ///     构造某一行的单元格枚举器
        /// </summary>
        /// <param name="row"></param>
        public EnumerableCell(IRow row)
        {
            _row = row;
        }
        public IEnumerator<ICell> GetEnumerator()
        {
            if (_row.PhysicalNumberOfCells == 0)
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