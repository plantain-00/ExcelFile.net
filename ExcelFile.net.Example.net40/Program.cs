using NPOI.HSSF.Util;

namespace ExcelFile.net.Example.net40
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var excel = new ExcelFile();
            excel.Sheet("test sheet");
            excel.Row().Cell("test1").Cell(2);
            excel.Row().Cell("test2").Cell(3);
            excel.Save("a.xls");
            var excel2 = new ExcelFile();
            excel2.Sheet("test2 sheet");
            excel2.Row(25, excel2.NewStyle().Background(HSSFColor.Yellow.Index)).Empty(2).Cell("test1");
            excel2.Row(15).Empty().Cell(1).Cell(2, excel2.NewStyle().Color(HSSFColor.Red.Index));
            excel2.Save("b.xls");
        }
    }
}