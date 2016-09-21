using System;

using NPOI.SS.UserModel;

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
        /// <param name="cell">the errored cell</param>
        public ExcelDataException(ICell cell)
        {
            Cell = cell;
        }

        /// <summary>
        ///     Cosntruct an ExcelDataException object
        /// </summary>
        /// <param name="message">the error message</param>
        /// <param name="cell">the errored cell</param>
        public ExcelDataException(string message, ICell cell) : base(message)
        {
            Cell = cell;
        }

        /// <summary>
        ///     Cosntruct an ExcelDataException object
        /// </summary>
        /// <param name="message">the error message</param>
        /// <param name="innerException">the inner exception</param>
        /// <param name="cell">the errored cell</param>
        public ExcelDataException(string message, Exception innerException, ICell cell) : base(message, innerException)
        {
            Cell = cell;
        }

        /// <summary>
        ///     The errored cell
        /// </summary>
        public ICell Cell { get; set; }
    }
}