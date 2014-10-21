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
    ///     提供Excel文件输入输出方法
    /// </summary>
    public static class ExcelUtils
    {
        /// <summary>
        ///     读取Excel文件
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
        ///     新建空Excel工作簿
        /// </summary>
        /// <param name="is2007OrLater"></param>
        /// <returns></returns>
        public static IWorkbook New(bool is2007OrLater = false)
        {
            return is2007OrLater ? new XSSFWorkbook() as IWorkbook : new HSSFWorkbook();
        }

        /// <summary>
        ///     远程下载Excel文件
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName">带扩展名</param>
        /// <param name="workbook"></param>
        public static void Save(IWorkbook workbook, HttpResponse response, string fileName)
        {
            response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(fileName, Encoding.UTF8));
            workbook.Write(response.OutputStream);
        }

        /// <summary>
        ///     本地保存Excel文件
        /// </summary>
        /// <param name="file">带扩展名</param>
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
        ///     远程下载Excel文件
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName">带扩展名</param>
        /// <param name="workbook"></param>
        public static void Save(IWorkbook workbook, HttpResponseBase response, string fileName)
        {
            response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(fileName, Encoding.UTF8));
            workbook.Write(response.OutputStream);
        }

        /// <summary>
        ///     取字符串
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
        ///     取数字
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
        ///     取Boolean
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
        ///     取日期
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
#endif
    }
}