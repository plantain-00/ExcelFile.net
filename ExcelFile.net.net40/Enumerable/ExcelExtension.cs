using System.Collections.Generic;

using NPOI.SS.UserModel;

namespace ExcelFile.net.Enumerable
{
    /// <summary>
    ///     提供一组用于查询实现 IEnumerable&lt;T> 的对象的 static（在 Visual Basic 中为 Shared）方法。
    /// </summary>
    public static class ExcelExtension
    {
        /// <summary>
        ///     返回类型为 IEnumerable&lt;T> 的输入。
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static IEnumerable<ISheet> AsEnumerable(this IWorkbook workbook)
        {
            return new EnumerableSheet(workbook);
        }
        /// <summary>
        ///     返回类型为 IEnumerable&lt;T> 的输入。
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static IEnumerable<IRow> AsEnumerable(this ISheet sheet)
        {
            return new EnumerableRow(sheet);
        }
        /// <summary>
        ///     返回类型为 IEnumerable&lt;T> 的输入。
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public static IEnumerable<ICell> AsEnumerable(this IRow row)
        {
            return new EnumerableCell(row);
        }
    }
}