using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

using System.Web;

using NPOI.SS.UserModel;
#if !NET20 &&!NET30 &&!NET35
using ExcelFile.net.Enumerable;
using System.Threading.Tasks;
#endif

namespace ExcelFile.net
{
    /// <summary>
    ///     An interface to render an Excel template with data
    /// </summary>
    public interface IExcelEditor
    {
        /// <summary>
        ///     Start mark of a template variable
        /// </summary>
        string StartMark { get; set; }

        /// <summary>
        ///     End mark of a template variable
        /// </summary>
        string EndMark { get; set; }

        /// <summary>
        ///     Separator between a template variable name and its property or field 
        /// </summary>
        string Separator { get; set; }

        /// <summary>
        ///     Warning messages created when rendering the template
        /// </summary>
        List<string> WarningMessages { get; }

        /// <summary>
        ///     Set the value of a cell
        /// </summary>
        /// <param name="name">the name of the cell</param>
        /// <param name="value">the value of the cell</param>
        void Set(string name, string value);

        /// <summary>
        ///     Set the value of a cell
        /// </summary>
        /// <param name="name">the name of the cell</param>
        /// <param name="value">the value of the cell</param>
        /// <param name="format">the format of the value</param>
        void Set(string name, double value, string format = null);

        /// <summary>
        ///     Set the value of a cell
        /// </summary>
        /// <param name="name">the name of the cell</param>
        /// <param name="value">the value of the cell</param>
        void Set(string name, bool value);

        /// <summary>
        ///     Set the value of a cell
        /// </summary>
        /// <param name="name">the name of the cell</param>
        /// <param name="value">the value of the cell</param>
        /// <param name="format">the format of the value</param>
        void Set(string name, DateTime value, string format = null);

        /// <summary>
        ///     Set the values of cells, the result is multiple rows
        /// </summary>
        /// <param name="name">the name of the cell</param>
        /// <param name="values">the values of the cell</param>
        /// <param name="willCopyRow">if it is true, the rows created by copy from the template row, or just create a new blank row</param>
        /// <param name="type">the type iof T</param>
        void Set<T>(string name, IList<T> values, bool willCopyRow = true, Type type = null);

        /// <summary>
        ///     Download the Excel file, for asp.net MVC, can use `return new EmptyResult();` as the response.
        /// </summary>
        /// <param name="response">the HTTP response</param>
        /// <param name="fileName">the file name</param>
        void Save(HttpResponse response, string fileName);

        /// <summary>
        ///     Save the file as a local file
        /// </summary>
        /// <param name="filePath">the target file path</param>
        void Save(string filePath);

#if !NET20 &&!NET30 &&!NET35
        /// <summary>
        ///     Download the Excel file, for asp.net MVC, can use `return new EmptyResult();` as the response.
        /// </summary>
        /// <param name="response">the HTTP response</param>
        /// <param name="fileName">the file name</param>
        void Save(HttpResponseBase response, string fileName);

        /// <summary>
        ///     Update all cell value if it is a formula
        /// </summary>
        void UpdateFormula();

        /// <summary>
        ///     Download the Excel file, for asp.net MVC, can use `return new EmptyResult();` as the response.
        /// </summary>
        /// <param name="response">the HTTP response</param>
        /// <param name="fileName">the file name</param>
        Task SaveAsync(HttpResponse response, string fileName);

        /// <summary>
        ///     Save the file as a local file
        /// </summary>
        /// <param name="filePath">the target file path</param>
        Task SaveAsync(string filePath);

        /// <summary>
        ///     Download the Excel file, for asp.net MVC, can use `return new EmptyResult();` as the response.
        /// </summary>
        /// <param name="response">the HTTP response</param>
        /// <param name="fileName">the file name</param>
        Task SaveAsync(HttpResponseBase response, string fileName);
#endif
    }

    /// <summary>
    ///     An class to render an Excel template with data
    /// </summary>
    public class ExcelEditor : IExcelEditor
    {
        /// <summary>
        ///     Current workbook
        /// </summary>
        public readonly IWorkbook Workbook;

        /// <summary>
        ///     Construct an ExcelEditor object
        /// </summary>
        /// <param name="file">the file path</param>
        /// <param name="is2007OrLater">whether the format of the excel file is 2007 or later</param>
        /// <param name="willJudgeByExtensionName">whether the format of the excel file will be judged by its extension name</param>
        public ExcelEditor(string file, bool is2007OrLater = false, bool willJudgeByExtensionName = true)
        {
            Workbook = ExcelUtils.New(file, FileMode.Open, FileAccess.Read, is2007OrLater, willJudgeByExtensionName);
            WarningMessages = new List<string>();
            StartMark = "{";
            EndMark = "}";
            Separator = ".";
        }

        /// <summary>
        ///     Start mark of a template variable
        /// </summary>
        public string StartMark { get; set; }

        /// <summary>
        ///     End mark of a template variable
        /// </summary>
        public string EndMark { get; set; }

        /// <summary>
        ///     Separator between a template variable name and its property or field 
        /// </summary>
        public string Separator { get; set; }

        /// <summary>
        ///     Warning messages created when rendering the template
        /// </summary>
        public List<string> WarningMessages { get; private set; }

        private void AddWarningMessage(string message)
        {
            foreach (var warningMessage in WarningMessages)
            {
                if (warningMessage == message)
                {
                    return;
                }
            }
            WarningMessages.Add(message);
        }

        /// <summary>
        ///     Set the value of a cell
        /// </summary>
        /// <param name="name">the name of the cell</param>
        /// <param name="value">the value of the cell</param>
        public void Set(string name, string value)
        {
            var placeHolderName = GetPlaceHolderName(name);
            foreach (var cell in FindCells(placeHolderName))
            {
                cell.SetCellValue(cell.StringCellValue.Replace(placeHolderName, value));
            }
        }

        /// <summary>
        ///     Set the value of a cell
        /// </summary>
        /// <param name="name">the name of the cell</param>
        /// <param name="value">the value of the cell</param>
        /// <param name="format">the format of the value</param>
        public void Set(string name, double value, string format = null)
        {
            var placeHolderName = GetPlaceHolderName(name);
            foreach (var cell in FindCells(placeHolderName))
            {
                if (cell.StringCellValue == placeHolderName)
                {
                    cell.SetCellValue(value);
                }
                else
                {
                    cell.SetCellValue(format == null ? cell.StringCellValue.Replace(placeHolderName, value.ToString()) : cell.StringCellValue.Replace(name, value.ToString(format)));
                }
            }
        }

        /// <summary>
        ///     Set the value of a cell
        /// </summary>
        /// <param name="name">the name of the cell</param>
        /// <param name="value">the value of the cell</param>
        public void Set(string name, bool value)
        {
            var placeHolderName = GetPlaceHolderName(name);
            foreach (var cell in FindCells(placeHolderName))
            {
                if (cell.StringCellValue == placeHolderName)
                {
                    cell.SetCellValue(value);
                }
                else
                {
                    cell.SetCellValue(cell.StringCellValue.Replace(placeHolderName, value.ToString()));
                }
            }
        }

        /// <summary>
        ///     Set the value of a cell
        /// </summary>
        /// <param name="name">the name of the cell</param>
        /// <param name="value">the value of the cell</param>
        /// <param name="format">the format of the value</param>
        public void Set(string name, DateTime value, string format = null)
        {
            var placeHolderName = GetPlaceHolderName(name);
            foreach (var cell in FindCells(placeHolderName))
            {
                if (cell.StringCellValue == placeHolderName)
                {
                    cell.SetCellValue(value);
                }
                else
                {
                    cell.SetCellValue(format == null ? cell.StringCellValue.Replace(placeHolderName, value.ToString()) : cell.StringCellValue.Replace(name, value.ToString(format)));
                }
            }
        }

        private void Set<T>(ICell cell, string name, T value, string cellValue, Type type = null, string format = null)
        {
            if (value == null)
            {
                cell.SetCellType(CellType.Blank);
                return;
            }
            if (type == null)
            {
                type = value.GetType();
            }
            if (type == typeof(string))
            {
                cell.SetCellValue(cellValue.Replace(name, value as string));
            }
            else if (type == typeof(DateTime))
            {
                if (cellValue == name)
                {
                    cell.SetCellValue(Convert.ToDateTime(value));
                }
                else
                {
                    cell.SetCellValue(format == null ? cellValue.Replace(name, value.ToString()) : cellValue.Replace(name, Convert.ToDateTime(value).ToString(format)));
                }
            }
            else if (value is DateTime?)
            {
                var dateTime = value as DateTime?;
                if (cellValue == name)
                {
                    cell.SetCellValue(dateTime.Value);
                }
                else
                {
                    cell.SetCellValue(format == null ? cellValue.Replace(name, dateTime.Value.ToString()) : cellValue.Replace(name, dateTime.Value.ToString(format)));
                }
            }
            else if (type == typeof(bool))
            {
                if (cellValue == name)
                {
                    cell.SetCellValue(Convert.ToBoolean(value));
                }
                else
                {
                    cell.SetCellValue(cellValue.Replace(name, value.ToString()));
                }
            }
            else if (value is bool?)
            {
                var boolean = value as bool?;
                if (cellValue == name)
                {
                    cell.SetCellValue(boolean.Value);
                }
                else
                {
                    cell.SetCellValue(cellValue.Replace(name, boolean.Value.ToString()));
                }
            }
            else if (type == typeof(double)
                     || type == typeof(float)
                     || type == typeof(int)
                     || type == typeof(uint)
                     || type == typeof(short)
                     || type == typeof(long)
                     || type == typeof(ushort)
                     || type == typeof(ulong)
                     || type == typeof(decimal))
            {
                SetDoubleValue(cell, name, cellValue, format, Convert.ToDouble(value));
            }
            else if (value is double?)
            {
                var d = value as double?;
                SetDoubleValue(cell, name, cellValue, format, d.Value);
            }
            else if (value is float?)
            {
                var d = value as float?;
                SetDoubleValue(cell, name, cellValue, format, d.Value);
            }
            else if (value is int?)
            {
                var d = value as int?;
                SetDoubleValue(cell, name, cellValue, format, d.Value);
            }
            else if (value is uint?)
            {
                var d = value as uint?;
                SetDoubleValue(cell, name, cellValue, format, d.Value);
            }
            else if (value is short?)
            {
                var d = value as short?;
                SetDoubleValue(cell, name, cellValue, format, d.Value);
            }
            else if (value is long?)
            {
                var d = value as long?;
                SetDoubleValue(cell, name, cellValue, format, d.Value);
            }
            else if (value is ushort?)
            {
                var d = value as ushort?;
                SetDoubleValue(cell, name, cellValue, format, d.Value);
            }
            else if (value is ulong?)
            {
                var d = value as ulong?;
                SetDoubleValue(cell, name, cellValue, format, d.Value);
            }
            else if (value is decimal?)
            {
                var d = value as decimal?;
                SetDoubleValue(cell, name, cellValue, format, (double)d.Value);
            }
            else
            {
                cell.SetCellValue(cellValue.Replace(name, value.ToString()));
            }
        }

        private static void SetDoubleValue(ICell cell, string name, string cellValue, string format, double doubleValue)
        {
            if (cellValue == name)
            {
                cell.SetCellValue(doubleValue);
            }
            else
            {
                cell.SetCellValue(format == null ? cellValue.Replace(name, doubleValue.ToString()) : cellValue.Replace(name, doubleValue.ToString(format)));
            }
        }

        /// <summary>
        ///     Set the values of cells, the result is multiple rows
        /// </summary>
        /// <param name="name">the name of the cell</param>
        /// <param name="values">the values of the cell</param>
        /// <param name="willCopyRow">if it is true, the rows created by copy from the template row, or just create a new blank row</param>
        /// <param name="type">the type iof T</param>
        public void Set<T>(string name, IList<T> values, bool willCopyRow = true, Type type = null)
        {
            if (values == null)
            {
                throw new ArgumentNullException(nameof(values));
            }
            if (type == null)
            {
                type = typeof(T);
            }
            var properties = type.GetProperties();
            var fields = type.GetFields();
            if (properties.Length + fields.Length == 0)
            {
                return;
            }
            var startOfPlaceHolderName = GetStartOfPlaceHolderName(name);
            var row = FindRow(startOfPlaceHolderName);
            if (row == null)
            {
                return;
            }

            var cache = new Dictionary<string, MemberInfo>();
            foreach (var propertyInfo in properties)
            {
                var placeHolderName = GetPlaceHolderName(Combine(name, propertyInfo.Name));
                var cell = Find(row, placeHolderName);
                if (cell != null)
                {
                    cache.Add(propertyInfo.Name, new MemberInfo(placeHolderName, cell.ColumnIndex, propertyInfo.PropertyType, cell.StringCellValue));
                }
                else
                {
                    cache.Add(propertyInfo.Name, null);
                }
            }
            foreach (var fieldInfo in fields)
            {
                var placeHolderName = GetPlaceHolderName(Combine(name, fieldInfo.Name));
                var cell = Find(row, placeHolderName);
                if (cell != null)
                {
                    cache.Add(fieldInfo.Name, new MemberInfo(placeHolderName, cell.ColumnIndex, fieldInfo.FieldType, cell.StringCellValue));
                }
                else
                {
                    cache.Add(fieldInfo.Name, null);
                }
            }

            if (values.Count == 0)
            {
                if (willCopyRow)
                {
                    for (var i = row.FirstCellNum; i < row.LastCellNum; i++)
                    {
                        row.RemoveCell(row.GetCell(i));
                    }
                }
                else
                {
                    RemovePlaceHolder(row, startOfPlaceHolderName);
                }
                return;
            }
            if (willCopyRow)
            {
                for (var i = 0; i < values.Count - 1; i++)
                {
                    row.CopyRowTo(row.RowNum + 1 + i);
                }
            }
            for (var i = 0; i < values.Count; i++)
            {
                var value = values[i];
                var nextRow = row.Sheet.GetRow(row.RowNum + 1);
                if (nextRow == null)
                {
                    nextRow = row.Sheet.CreateRow(row.RowNum + 1);
                    nextRow.RowStyle = row.RowStyle;
                }
                foreach (var propertyInfo in properties)
                {
                    var result = type.InvokeMember(propertyInfo.Name, BindingFlags.GetProperty, null, value, null);
                    var memberInfo = cache[propertyInfo.Name];
                    if (memberInfo == null)
                    {
                        break;
                    }
                    var placeHolderName = memberInfo.PlaceHolderName;
                    var cell = row.GetCell(memberInfo.ColumnIndex);
                    if (!willCopyRow
                        && i != values.Count - 1)
                    {
                        var nextCell = nextRow.GetCell(cell.ColumnIndex);
                        if (nextCell == null)
                        {
                            var newCell = nextRow.CreateCell(cell.ColumnIndex);
                            newCell.CellStyle = cell.CellStyle;
                        }
                    }
                    Set(cell, placeHolderName, result, memberInfo.Value, memberInfo.Type);
                }
                foreach (var fieldInfo in fields)
                {
                    var result = type.InvokeMember(fieldInfo.Name, BindingFlags.GetField, null, value, null);
                    var memberInfo = cache[fieldInfo.Name];
                    if (memberInfo == null)
                    {
                        break;
                    }
                    var placeHolderName = memberInfo.PlaceHolderName;
                    var cell = row.GetCell(memberInfo.ColumnIndex);
                    if (!willCopyRow
                        && i != values.Count - 1)
                    {
                        var nextCell = nextRow.GetCell(cell.ColumnIndex);
                        if (nextCell == null)
                        {
                            var newCell = nextRow.CreateCell(cell.ColumnIndex);
                            newCell.CellStyle = cell.CellStyle;
                        }
                    }
                    Set(cell, placeHolderName, result, memberInfo.Value, memberInfo.Type);
                }
                row = nextRow;
            }
        }

        private IEnumerable<ICell> FindCells(string name)
        {
            var result = new List<ICell>();
            for (var i = 0; i < Workbook.NumberOfSheets; i++)
            {
                var sheet = Workbook.GetSheetAt(i);
                for (var j = sheet.FirstRowNum; j <= sheet.LastRowNum; j++)
                {
                    var row = sheet.GetRow(j);
                    if (row == null)
                    {
                        row = sheet.CreateRow(j);
                    }
                    for (var k = row.FirstCellNum; k < row.LastCellNum; k++)
                    {
                        var cell = row.GetCell(k);
                        if (cell == null)
                        {
                            cell = row.CreateCell(k);
                        }
                        if (cell.CellType == CellType.String
                            && cell.StringCellValue.Contains(name))
                        {
                            result.Add(cell);
                        }
                    }
                }
            }
            if (result.Count == 0)
            {
                AddWarningMessage($"variable \"{name}\" is not used here, so it can be deleted.");
            }
            return result;
        }

        private IRow FindRow(string name)
        {
            for (var i = 0; i < Workbook.NumberOfSheets; i++)
            {
                var sheet = Workbook.GetSheetAt(i);
                for (var j = sheet.FirstRowNum; j <= sheet.LastRowNum; j++)
                {
                    var row = sheet.GetRow(j);
                    if (row == null)
                    {
                        row = sheet.CreateRow(j);
                    }
                    for (var k = row.FirstCellNum; k < row.LastCellNum; k++)
                    {
                        var cell = row.GetCell(k);
                        if (cell == null)
                        {
                            cell = row.CreateCell(k);
                        }
                        if (cell.CellType == CellType.String
                            && cell.StringCellValue.Contains(name))
                        {
                            return cell.Row;
                        }
                    }
                }
            }
            AddWarningMessage($"variable \"{name}\" is not used here, so it can be deleted.");
            return null;
        }

        private static void RemovePlaceHolder(IRow row, string name)
        {
            for (var k = 0; k < row.LastCellNum; k++)
            {
                var cell = row.GetCell(k);
                if (cell == null)
                {
                    cell = row.CreateCell(k);
                }
                if (cell.CellType == CellType.String
                    && cell.StringCellValue.Contains(name))
                {
                    row.RemoveCell(cell);
                }
            }
        }

        private ICell Find(IRow row, string name)
        {
            for (var k = 0; k < row.PhysicalNumberOfCells; k++)
            {
                var cell = row.Cells[k];
                if (cell.CellType == CellType.String
                    && cell.StringCellValue.Contains(name))
                {
                    return cell;
                }
            }
            AddWarningMessage(name + " not found in row:" + row.RowNum);
            return null;
        }

        /// <summary>
        ///     "VariableA"->"{VariableA}"
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private string GetPlaceHolderName(string name)
        {
            if (name.Contains(StartMark)
                || name.Contains(EndMark))
            {
                AddWarningMessage($"the name of variable \"{name}\" should include the start mark or end mark.");
            }
            return string.Format("{1}{0}{2}", name, StartMark, EndMark);
        }

        /// <summary>
        ///     "ClassA"、"MemberB"->"ClassA-MemberB"
        /// </summary>
        /// <param name="name"></param>
        /// <param name="memberName"></param>
        /// <returns></returns>
        private string Combine(string name, string memberName)
        {
            return name + Separator + memberName;
        }

        /// <summary>
        ///     "ClassA"->"{ClassA-"
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private string GetStartOfPlaceHolderName(string name)
        {
            if (name.Contains(StartMark)
                || name.Contains(EndMark)
                || name.Contains(Separator))
            {
                AddWarningMessage($"the name of variable \"{name}\" should not include start mark, end mark and separator.");
            }
            return string.Format("{1}{0}{2}", name, StartMark, Separator);
        }

        /// <summary>
        ///     Download the Excel file, for asp.net MVC, can use `return new EmptyResult();` as the response.
        /// </summary>
        /// <param name="response">the HTTP response</param>
        /// <param name="fileName">the file name</param>
        public void Save(HttpResponse response, string fileName)
        {
            ExcelUtils.Save(Workbook, response, fileName);
        }

        /// <summary>
        ///     Save the file as a local file
        /// </summary>
        /// <param name="filePath">the target file path</param>
        public void Save(string filePath)
        {
            ExcelUtils.Save(Workbook, filePath);
        }

#if !NET20 &&!NET30 &&!NET35
        /// <summary>
        ///     Download the Excel file, for asp.net MVC, can use `return new EmptyResult();` as the response.
        /// </summary>
        /// <param name="response">the HTTP response</param>
        /// <param name="fileName">the file name</param>
        public void Save(HttpResponseBase response, string fileName)
        {
            ExcelUtils.Save(Workbook, response, fileName);
        }

        /// <summary>
        ///     Update all cell value if it is a formula
        /// </summary>
        public void UpdateFormula()
        {
            var formulaEvaluator = Workbook.GetFormulaEvaluator();
            foreach (var sheet in Workbook.AsEnumerable())
            {
                sheet.ForceFormulaRecalculation = true;
                foreach (var row in sheet.AsEnumerable())
                {
                    foreach (var cell in row.AsEnumerable())
                    {
                        if (cell != null
                            && cell.CellType == CellType.Formula)
                        {
                            formulaEvaluator.EvaluateFormulaCell(cell);
                        }
                    }
                }
            }
        }

        /// <summary>
        ///     Download the Excel file, for asp.net MVC, can use `return new EmptyResult();` as the response.
        /// </summary>
        /// <param name="response">the HTTP response</param>
        /// <param name="fileName">the file name</param>
        public Task SaveAsync(HttpResponse response, string fileName)
        {
            return Task.Factory.StartNew(() => Save(response, fileName));
        }

        /// <summary>
        ///     Save the file as a local file
        /// </summary>
        /// <param name="filePath">the target file path</param>
        public Task SaveAsync(string filePath)
        {
            return Task.Factory.StartNew(() => Save(filePath));
        }

        /// <summary>
        ///     Download the Excel file, for asp.net MVC, can use `return new EmptyResult();` as the response.
        /// </summary>
        /// <param name="response">the HTTP response</param>
        /// <param name="fileName">the file name</param>
        public Task SaveAsync(HttpResponseBase response, string fileName)
        {
            return Task.Factory.StartNew(() => Save(response, fileName));
        }
#endif
    }

    internal class MemberInfo
    {
        public MemberInfo(string placeHolderName, int columnIndex, Type type, string value)
        {
            PlaceHolderName = placeHolderName;
            ColumnIndex = columnIndex;
            Type = type;
            Value = value;
        }

        public string PlaceHolderName { get; set; }
        public int ColumnIndex { get; set; }
        public Type Type { get; set; }
        public string Value { get; set; }
    }
}