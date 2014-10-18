using System;

namespace ExcelFile.net
{
    /// <summary>
    ///     Excel数据异常
    /// </summary>
    public class ExcelDataException : Exception
    {
        /// <summary>
        ///     构造ExcelDataException对象
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        public ExcelDataException(int rowIndex = 0, int columnIndex = 0)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        /// <summary>
        ///     构造ExcelDataException对象
        /// </summary>
        /// <param name="message"></param>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        public ExcelDataException(string message, int rowIndex = 0, int columnIndex = 0) : base(message)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        /// <summary>
        ///     构造ExcelDataException对象
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        public ExcelDataException(string message, Exception innerException, int rowIndex = 0, int columnIndex = 0) : base(message, innerException)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        /// <summary>
        ///     行索引
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        ///     列索引
        /// </summary>
        public int ColumnIndex { get; set; }
    }
}