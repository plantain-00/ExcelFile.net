using System.Collections;
using System.Collections.Generic;

using NPOI.SS.UserModel;

namespace ExcelFile.net.Enumerable
{
    /// <summary>
    ///     某个工作簿的工作表的枚举器
    /// </summary>
    public class EnumerableSheet : IEnumerable<ISheet>
    {
        private readonly IWorkbook _workbook;
        /// <summary>
        ///     构造某个工作簿的工作表的枚举器
        /// </summary>
        /// <param name="workbook"></param>
        public EnumerableSheet(IWorkbook workbook)
        {
            _workbook = workbook;
        }
        public IEnumerator<ISheet> GetEnumerator()
        {
            for (var i = 0; i < _workbook.NumberOfSheets; i++)
            {
                yield return _workbook.GetSheetAt(i);
            }
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}