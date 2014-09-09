ExcelFile.net
=============

+ a Fluent Excel File Writer based on NPOI.
+ a Excel template editor.
+ a enumerator of worksheets, rows and cells.

        var excel = new ExcelFile();
    excel.Sheet("test sheet");
    excel.Row().Cell("test1").Cell(2);
    excel.Row().Cell("test2").Cell(3);
    excel.Save("a.xls");

![](https://raw.githubusercontent.com/plantain-00/ExcelFile.net/master/images/a.JPG)

	var excel2 = new ExcelFile();
    excel2.Sheet("test2 sheet");
    excel2.Row(25, excel2.NewStyle().Background(HSSFColor.Yellow.Index)).Empty(2).Cell("test1");
    excel2.Row(15).Empty().Cell(1).Cell(2, excel2.NewStyle().Color(HSSFColor.Red.Index));
    excel2.Save("b.xls");

![](https://raw.githubusercontent.com/plantain-00/ExcelFile.net/master/images/b.JPG)

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
	excel.Save("d.xls");

![](https://raw.githubusercontent.com/plantain-00/ExcelFile.net/master/images/c.PNG)
![](https://raw.githubusercontent.com/plantain-00/ExcelFile.net/master/images/d.PNG)

## reference
### ExcelFile

+ 内容：工作表Sheet()、行Row()、单元格Cell()、空的单元格Empty()、合并单元格Cell()
+ 单元格样式：默认样式Style、新样式NewStyle()、内联样式Cell()、行样式Row()
+ 列样式：列宽Sheet()
+ 行样式：默认行高DefaultRowHeight()、内联行高Row()
+ 输出：本地文件Save()、远程下载Save()

### ExcelStyle

+ 背景色：Background
+ 边框及边框颜色：Border、BorderTop、BorderBottom、BorderLeft、BorderRight
+ 对齐：Align、VerticalAlign
+ 文字：WrapText、Italic、Underline、FontSize、Font、Color、Bold

## nuget
You can get [it](https://www.nuget.org/packages/ExcelFile.net) from Nuget.
