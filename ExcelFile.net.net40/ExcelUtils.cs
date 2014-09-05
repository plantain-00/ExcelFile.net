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
        /// <param name="is2007OrMore"></param>
        /// <returns></returns>
        public static IWorkbook New(string file, FileMode fileMode, FileAccess fileAccess, bool is2007OrMore = false)
        {
            using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                return is2007OrMore ? new XSSFWorkbook(stream) as IWorkbook : new HSSFWorkbook(stream);
            }
        }

        /// <summary>
        ///     新建空Excel工作簿
        /// </summary>
        /// <param name="is2007OrMore"></param>
        /// <returns></returns>
        public static IWorkbook New(bool is2007OrMore = false)
        {
            return is2007OrMore ? new XSSFWorkbook() as IWorkbook : new HSSFWorkbook();
        }

        /// <summary>
        ///     远程下载Excel文件
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName">带扩展名</param>
        /// <param name="workbook"></param>
        public static void Save(IWorkbook workbook, HttpResponse response, string fileName)
        {
            response.ContentType = "application/vnd.ms-excel";
            response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(fileName, Encoding.UTF8));
            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                stream.Flush();
                stream.Position = 0;
                stream.WriteTo(response.OutputStream);
            }
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
            response.ContentType = "application/vnd.ms-excel";
            response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(fileName, Encoding.UTF8));
            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                stream.Flush();
                stream.Position = 0;
                stream.WriteTo(response.OutputStream);
            }
        }
#endif
    }
}