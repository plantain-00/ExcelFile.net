using System.Collections;
using System.Collections.Generic;

using NPOI.SS.UserModel;

namespace ExcelFile.net.Enumerable
{
    /// <summary>
    ///     某个工作表的行的枚举器
    /// </summary>
    public class EnumerableRow : IEnumerable<IRow>
    {
        private readonly ISheet _sheet;
        /// <summary>
        ///     构造某个工作表的行的枚举器
        /// </summary>
        /// <param name="sheet"></param>
        public EnumerableRow(ISheet sheet)
        {
            _sheet = sheet;
        }
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