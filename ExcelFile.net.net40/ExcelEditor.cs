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
    ///     向Excel模板中插入数据的接口
    /// </summary>
    public interface IExcelEditor
    {
        /// <summary>
        ///     开始标记
        /// </summary>
        string StartMark { get; set; }

        /// <summary>
        ///     结束标记
        /// </summary>
        string EndMark { get; set; }

        /// <summary>
        ///     分隔符
        /// </summary>
        string Separator { get; set; }

        /// <summary>
        ///     警告信息
        /// </summary>
        List<string> WarningMessage { get; }

        /// <summary>
        ///     设置单元格的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        void Set(string name, string value);

        /// <summary>
        ///     设置单元格的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="format"></param>
        void Set(string name, double value, string format = null);

        /// <summary>
        ///     设置单元格的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        void Set(string name, bool value);

        /// <summary>
        ///     设置单元格的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="format"></param>
        void Set(string name, DateTime value, string format = null);

        /// <summary>
        ///     设置单元格的值
        ///     在外部缓存T的Type会提高性能
        /// </summary>
        /// <param name="name"></param>
        /// <param name="values"></param>
        /// <param name="willCopyRow"></param>
        /// <param name="type"></param>
        void Set<T>(string name, IList<T> values, bool willCopyRow = true, Type type = null);

        /// <summary>
        ///     远程下载Excel文件，MVC中return new EmptyResult();
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName"></param>
        void Save(HttpResponse response, string fileName);

        /// <summary>
        ///     本地保存Excel文件
        /// </summary>
        /// <param name="file"></param>
        void Save(string file);

#if !NET20 &&!NET30 &&!NET35
        /// <summary>
        ///     远程下载Excel文件，MVC中return new EmptyResult();
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName"></param>
        void Save(HttpResponseBase response, string fileName);

        /// <summary>
        ///     更新Formula
        /// </summary>
        void UpdateFormula();

        /// <summary>
        ///     远程下载Excel文件，MVC中return new EmptyResult();
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName"></param>
        Task SaveAsync(HttpResponse response, string fileName);

        /// <summary>
        ///     本地保存Excel文件
        /// </summary>
        /// <param name="file"></param>
        Task SaveAsync(string file);

        /// <summary>
        ///     远程下载Excel文件，MVC中return new EmptyResult();
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName"></param>
        Task SaveAsync(HttpResponseBase response, string fileName);
#endif
    }

    /// <summary>
    ///     向Excel模板中插入数据的类
    /// </summary>
    public class ExcelEditor : IExcelEditor
    {
        /// <summary>
        ///     当前工作簿
        /// </summary>
        public readonly IWorkbook Workbook;

        /// <summary>
        ///     构造Excel编辑对象
        /// </summary>
        /// <param name="file"></param>
        /// <param name="is2007OrLater"></param>
        /// <param name="willJudgeByExtensionName"></param>
        public ExcelEditor(string file, bool is2007OrLater = false, bool willJudgeByExtensionName = true)
        {
            Workbook = ExcelUtils.New(file, FileMode.Open, FileAccess.Read, is2007OrLater, willJudgeByExtensionName);
            WarningMessage = new List<string>();
            StartMark = "{";
            EndMark = "}";
            Separator = ".";
        }

        public string StartMark { get; set; }
        public string EndMark { get; set; }
        public string Separator { get; set; }

        /// <summary>
        ///     警告信息
        /// </summary>
        public List<string> WarningMessage { get; private set; }

        private void AddWarningMessage(string message)
        {
            foreach (var warningMessage in WarningMessage)
            {
                if (warningMessage == message)
                {
                    return;
                }
            }
            WarningMessage.Add(message);
        }

        /// <summary>
        ///     设置单元格的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public void Set(string name, string value)
        {
            var placeHolderName = GetPlaceHolderName(name);
            foreach (var cell in FindCells(placeHolderName))
            {
                cell.SetCellValue(cell.StringCellValue.Replace(placeHolderName, value));
            }
        }

        /// <summary>
        ///     设置单元格的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="format"></param>
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
        ///     设置单元格的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
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
        ///     设置单元格的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="format"></param>
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
            if (type == typeof (string))
            {
                cell.SetCellValue(cellValue.Replace(name, value as string));
            }
            else if (type == typeof (DateTime))
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
            else if (type == typeof (bool))
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
            else if (type == typeof (double)
                     || type == typeof (float)
                     || type == typeof (int)
                     || type == typeof (uint)
                     || type == typeof (short)
                     || type == typeof (long)
                     || type == typeof (ushort)
                     || type == typeof (ulong)
                     || type == typeof (decimal))
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
                SetDoubleValue(cell, name, cellValue, format, (double) d.Value);
            }
            else
            {
                AddWarningMessage("cannot support type:" + type.FullName);
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
        ///     设置单元格的值
        ///     在外部缓存T的Type会提高性能
        /// </summary>
        /// <param name="name"></param>
        /// <param name="values"></param>
        /// <param name="willCopyRow"></param>
        /// <param name="type"></param>
        public void Set<T>(string name, IList<T> values, bool willCopyRow = true, Type type = null)
        {
            if (values == null)
            {
                throw new ArgumentNullException("values");
            }
            if (type == null)
            {
                type = typeof (T);
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
                            nextRow.CreateCell(cell.ColumnIndex);
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
                            nextRow.CreateCell(cell.ColumnIndex);
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
                var sheet = Workbook[i];
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
                AddWarningMessage(string.Format("变量\"{0}\"未被使用，可以删除", name));
            }
            return result;
        }

        private IRow FindRow(string name)
        {
            for (var i = 0; i < Workbook.NumberOfSheets; i++)
            {
                var sheet = Workbook[i];
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
            AddWarningMessage(string.Format("变量\"{0}\"未被使用，可以删除", name));
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
                AddWarningMessage(string.Format("变量名\"{0}\"不应该包含大括号", name));
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
                AddWarningMessage(string.Format("变量名\"{0}\"不应该包含大括号或'-'", name));
            }
            return string.Format("{1}{0}{2}", name, StartMark, Separator);
        }

        /// <summary>
        ///     远程下载Excel文件，MVC中return new EmptyResult();
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName"></param>
        public void Save(HttpResponse response, string fileName)
        {
            ExcelUtils.Save(Workbook, response, fileName);
        }

        /// <summary>
        ///     本地保存Excel文件
        /// </summary>
        /// <param name="file"></param>
        public void Save(string file)
        {
            ExcelUtils.Save(Workbook, file);
        }

#if !NET20 &&!NET30 &&!NET35
        /// <summary>
        ///     远程下载Excel文件，MVC中return new EmptyResult();
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName"></param>
        public void Save(HttpResponseBase response, string fileName)
        {
            ExcelUtils.Save(Workbook, response, fileName);
        }

        /// <summary>
        ///     更新Formula
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

        public Task SaveAsync(HttpResponse response, string fileName)
        {
            return Task.Factory.StartNew(() => Save(response, fileName));
        }

        public Task SaveAsync(string file)
        {
            return Task.Factory.StartNew(() => Save(file));
        }

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