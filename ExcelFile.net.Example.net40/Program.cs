using System;
using System.IO;

using ExcelFile.net.Enumerable;

using NPOI.HSSF.Util;
using NPOI.SS.UserModel;

namespace ExcelFile.net.Example.net40
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            //Example A from A.xlsx
            IExcelEditor excelA = new ExcelEditor("../../A.xlsx");
            excelA.Set("name", "Sara");
            excelA.Set("age", 123);
            excelA.Save("../../A_result.xlsx");

            //Example B from B.xlsx
            IExcelEditor excelB = new ExcelEditor("../../B.xlsx");
            excelB.Set("s",
                       new[]
                       {
                           new
                           {
                               Name = "Tommy",
                               Age = 12
                           },
                           new
                           {
                               Name = "Philips",
                               Age = 13
                           }
                       });
            excelB.Save("../../B_result.xlsx");

            //Example C from C.xlsx
            IExcelEditor excelC = new ExcelEditor("../../C.xlsx");
            excelC.Set("s",
                       new[]
                       {
                           new
                           {
                               Name = "Tommy",
                               Age = 12
                           },
                           new
                           {
                               Name = "Philips",
                               Age = 13
                           }
                       },
                       false);
            excelC.Save("../../C_result.xlsx");

            //Example D
            var excelD = new ExcelFile();
            excelD.Sheet("D sheet");
            excelD.Row().Cell("Tommy").Cell(12);
            excelD.Row().Cell("Philips").Cell(13);
            excelD.Save("../../D_result.xls");

            //Example E
            var excelE = new ExcelFile(true);
            excelE.Sheet("E sheet");
            excelE.Row(25, excelE.NewStyle().Background(HSSFColor.Yellow.Index)).Empty(2).Cell("Tommy");
            excelE.Row(15).Empty().Cell(12).Cell(13, excelE.NewStyle().Color(HSSFColor.Red.Index));
            excelE.Save("../../E_result.xlsx");

            //Example F from F.xlsx
            foreach (var sheet in ExcelUtils.New("../../F.xlsx", FileMode.Open, FileAccess.Read).AsEnumerable())
            {
                foreach (var row in sheet.AsEnumerable())
                {
                    if (row == null)
                    {
                        continue;
                    }
                    foreach (var cell in row.AsEnumerable())
                    {
                        if (cell == null)
                        {
                            Console.WriteLine("null");
                        }
                        else switch (cell.CellType)
                        {
                            case CellType.Blank:
                                Console.WriteLine("blank");
                                break;
                            case CellType.Boolean:
                                Console.WriteLine(cell.BooleanCellValue);
                                break;
                            case CellType.Numeric:
                                Console.WriteLine(cell.NumericCellValue);
                                break;
                            case CellType.String:
                                Console.WriteLine(cell.StringCellValue);
                                break;
                        }
                    }
                }
            }
            Console.Read();
        }
    }
}