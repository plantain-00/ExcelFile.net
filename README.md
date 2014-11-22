ExcelFile.net
=============

+ a Excel template editor.
+ a Fluent Excel File Writer based on NPOI.
+ a enumerator of worksheets, rows and cells.


## Example A


            //Example A from A.xlsx
            IExcelEditor excelA = new ExcelEditor("../../A.xlsx");
            excelA.Set("name", "Sara");
            excelA.Set("age", 123);
            excelA.Save("../../A_result.xlsx");

![](https://raw.githubusercontent.com/plantain-00/ExcelFile.net/master/images/A.PNG)

![](https://raw.githubusercontent.com/plantain-00/ExcelFile.net/master/images/A_result.PNG)


## Example B


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

![](https://raw.githubusercontent.com/plantain-00/ExcelFile.net/master/images/B.PNG)

![](https://raw.githubusercontent.com/plantain-00/ExcelFile.net/master/images/B_result.PNG)


## Example C


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
            excelC.UpdateFormula();
            excelC.Save("../../C_result.xlsx");

![](https://raw.githubusercontent.com/plantain-00/ExcelFile.net/master/images/C.PNG)

![](https://raw.githubusercontent.com/plantain-00/ExcelFile.net/master/images/C_result.PNG)


## nuget
You can get [it](https://www.nuget.org/packages/ExcelFile.net) from Nuget.
