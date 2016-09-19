using System;
using System.IO;
using System.Text;
using System.Web;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelFile.net
{
    /// <summary>
    ///     I/O operations for Excel file
    /// </summary>
    public static class ExcelUtils
    {
        /// <summary>
        ///     Read an excel file
        /// </summary>
        /// <param name="file"></param>
        /// <param name="fileMode"></param>
        /// <param name="fileAccess"></param>
        /// <param name="is2007OrLater"></param>
        /// <param name="willJudgeByExtensionName"></param>
        /// <returns></returns>
        public static IWorkbook New(string file, FileMode fileMode, FileAccess fileAccess, bool is2007OrLater = false, bool willJudgeByExtensionName = true)
        {
            if (willJudgeByExtensionName)
            {
                if (file.EndsWith(".xls"))
                {
                    is2007OrLater = false;
                }
                else if (file.EndsWith(".xlsx"))
                {
                    is2007OrLater = true;
                }
            }
            using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                return is2007OrLater ? new XSSFWorkbook(stream) as IWorkbook : new HSSFWorkbook(stream);
            }
        }

        /// <summary>
        ///     Create a new Excel workbook
        /// </summary>
        /// <param name="is2007OrLater"></param>
        /// <returns></returns>
        public static IWorkbook New(bool is2007OrLater = false)
        {
            return is2007OrLater ? new XSSFWorkbook() as IWorkbook : new HSSFWorkbook();
        }

        /// <summary>
        ///     Read Excel workbook from a stream
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="is2007OrLater"></param>
        /// <returns></returns>
        public static IWorkbook New(Stream stream, bool is2007OrLater = false)
        {
            return is2007OrLater ? new XSSFWorkbook(stream) as IWorkbook : new HSSFWorkbook(stream);
        }

        /// <summary>
        ///     Download the Excel file, for asp.net MVC, can use `return new EmptyResult();` as the response.
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName">with extension name</param>
        /// <param name="workbook"></param>
        public static void Save(IWorkbook workbook, HttpResponse response, string fileName)
        {
            response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(fileName, Encoding.UTF8));
            workbook.Write(response.OutputStream);
        }

        /// <summary>
        ///     Save the file as a local file
        /// </summary>
        /// <param name="file">with extension name</param>
        /// <param name="workbook"></param>
        public static void Save(IWorkbook workbook, string file)
        {
            using (var stream = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }
        }

#if !NET20 &&!NET30 &&!NET35
        /// <summary>
        ///     Download the Excel file, for asp.net MVC, can use `return new EmptyResult();` as the response.
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName">with extension name</param>
        /// <param name="workbook"></param>
        public static void Save(IWorkbook workbook, HttpResponseBase response, string fileName)
        {
            response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(fileName, Encoding.UTF8));
            workbook.Write(response.OutputStream);
        }

        /// <summary>
        ///     Get string value from a cell
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="ExcelDataException"></exception>
        public static string GetString(this ICell cell)
        {
            try
            {
                return cell.StringCellValue;
            }
            catch (Exception exception)
            {
                throw new ExcelDataException("Error when get a string from a Excel's cell.", exception, cell.RowIndex, cell.ColumnIndex);
            }
        }

        /// <summary>
        ///     Get a number value from a cell
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="ExcelDataException"></exception>
        public static double GetNumber(this ICell cell)
        {
            try
            {
                return cell.NumericCellValue;
            }
            catch (Exception exception)
            {
                throw new ExcelDataException("Error when get a number from a Excel's cell.", exception, cell.RowIndex, cell.ColumnIndex);
            }
        }

        /// <summary>
        ///     Get a boolean value from a cell
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="ExcelDataException"></exception>
        public static bool GetBoolean(this ICell cell)
        {
            try
            {
                return cell.BooleanCellValue;
            }
            catch (Exception exception)
            {
                throw new ExcelDataException("Error when get a bool value from a Excel's cell.", exception, cell.RowIndex, cell.ColumnIndex);
            }
        }

        /// <summary>
        ///     Get a date value from a cell
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="ExcelDataException"></exception>
        public static DateTime GetDate(this ICell cell)
        {
            try
            {
                return cell.DateCellValue;
            }
            catch (Exception exception)
            {
                throw new ExcelDataException("Error when get a date from a Excel's cell.", exception, cell.RowIndex, cell.ColumnIndex);
            }
        }

        /// <summary>
        ///     Get FormulaEvaluator from a workbook
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static IFormulaEvaluator GetFormulaEvaluator(this IWorkbook workbook)
        {
            if (workbook is HSSFWorkbook)
            {
                return new HSSFFormulaEvaluator(workbook);
            }
            return new XSSFFormulaEvaluator(workbook);
        }
#endif
    }
}