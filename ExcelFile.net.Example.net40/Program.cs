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
        }
    }
}