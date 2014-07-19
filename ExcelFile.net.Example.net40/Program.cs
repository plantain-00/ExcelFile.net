using System;
using System.Collections.Generic;

namespace ExcelFile.net.Example.net40
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            //var excel = new ExcelFile();
            //excel.Sheet("test sheet");
            //excel.Row().Cell("test1").Cell(2);
            //excel.Row().Cell("test2").Cell(3);
            //excel.Save("a.xls");
            //var excel2 = new ExcelFile();
            //excel2.Sheet("test2 sheet");
            //excel2.Row(25, excel2.NewStyle().Background(HSSFColor.Yellow.Index)).Empty(2).Cell("test1");
            //excel2.Row(15).Empty().Cell(1).Cell(2, excel2.NewStyle().Color(HSSFColor.Red.Index));
            //excel2.Save("b.xls");
            var excel = new ExcelEditor("c.xls");
            excel.Set("测试", "sss");
            excel.Set("测试2", 123.456);
            excel.Set("测试3", false);
            excel.Set("测试4", DateTime.Now);
            var testData = new[]
                           {
                               new
                               {
                                   F1 = "aa",
                                   F2 = 12
                               },
                               new
                               {
                                   F1 = "bb",
                                   F2 = 121
                               }
                           };
            excel.Set("测试5", testData);
            excel.Set("测试6", testData, false);
            excel.Set("测试7", new List<ClassA>());
            excel.Set("测试8", new List<ClassA>(), false);
            excel.Save("d.xls");
        }
    }

    public class ClassA
    {
        public string F1 { get; set; }
        public int F2 { get; set; }
    }
}