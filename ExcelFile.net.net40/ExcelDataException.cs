using System;

namespace ExcelFile.net
{
    /// <summary>
    ///     Excel data exception class
    /// </summary>
    public class ExcelDataException : Exception
    {
        /// <summary>
        ///     Cosntruct an ExcelDataException object
        /// </summary>
        /// <param name="rowIndex">the row index of the errored cell</param>
        /// <param name="columnIndex">the column index of the errored cell</param>
        public ExcelDataException(int rowIndex = 0, int columnIndex = 0)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        /// <summary>
        ///     Cosntruct an ExcelDataException object
        /// </summary>
        /// <param name="message">the error message</param>
        /// <param name="rowIndex">the row index of the errored cell</param>
        /// <param name="columnIndex">the column index of the errored cell</param>
        public ExcelDataException(string message, int rowIndex = 0, int columnIndex = 0) : base(message)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        /// <summary>
        ///     Cosntruct an ExcelDataException object
        /// </summary>
        /// <param name="message">the error message</param>
        /// <param name="innerException">the inner exception</param>
        /// <param name="rowIndex">the row index of the errored cell</param>
        /// <param name="columnIndex">the column index of the errored cell</param>
        public ExcelDataException(string message, Exception innerException, int rowIndex = 0, int columnIndex = 0) : base(message, innerException)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        /// <summary>
        ///     The Row index of the errored cell
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        ///     The column index of the errored cell
        /// </summary>
        public int ColumnIndex { get; set; }
    }
}