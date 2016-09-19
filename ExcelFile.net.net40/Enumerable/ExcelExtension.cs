using System.Collections.Generic;

using NPOI.SS.UserModel;

namespace ExcelFile.net.Enumerable
{
    /// <summary>
    ///     Some extension methods that create enumerable sheets, rows or cells
    /// </summary>
    public static class ExcelExtension
    {
        /// <summary>
        ///     Create enumerable sheets from a workbook object
        /// </summary>
        /// <param name="workbook">the workbook object</param>
        /// <returns></returns>
        public static IEnumerable<ISheet> AsEnumerable(this IWorkbook workbook)
        {
            return new EnumerableSheet(workbook);
        }
        /// <summary>
        ///     Create enumerable rows from a sheet object
        /// </summary>
        /// <param name="sheet">the sheet object</param>
        /// <returns></returns>
        public static IEnumerable<IRow> AsEnumerable(this ISheet sheet)
        {
            return new EnumerableRow(sheet);
        }
        /// <summary>
        ///     Create enumerable cells from a row object
        /// </summary>
        /// <param name="row">the row object</param>
        /// <returns></returns>
        public static IEnumerable<ICell> AsEnumerable(this IRow row)
        {
            return new EnumerableCell(row);
        }
    }
}