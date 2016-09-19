using System.Collections;
using System.Collections.Generic;

using NPOI.SS.UserModel;

namespace ExcelFile.net.Enumerable
{
    /// <summary>
    ///     Enumerable sheets
    /// </summary>
    public class EnumerableSheet : IEnumerable<ISheet>
    {
        private readonly IWorkbook _workbook;
        /// <summary>
        ///     Construct an enumerable sheets from a workbook object
        /// </summary>
        /// <param name="workbook">the workbook object</param>
        public EnumerableSheet(IWorkbook workbook)
        {
            _workbook = workbook;
        }
        /// <summary>
        ///     The implementment of IEnumerable
        /// </summary>
        /// <returns></returns>
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