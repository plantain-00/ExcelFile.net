using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelFile.net
{
    /// <summary>
    ///     从Excel模板中插入数据的类
    /// </summary>
    public class ExcelEditor
    {
        private readonly IWorkbook _workbook;
        /// <summary>
        ///     构造Excel编辑对象
        /// </summary>
        /// <param name="file"></param>
        /// <param name="is2007OrMore"></param>
        public ExcelEditor(string file, bool is2007OrMore = false)
        {
            using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                _workbook = is2007OrMore ? new XSSFWorkbook(stream) as IWorkbook : new HSSFWorkbook(stream);
            }
        }
        /// <summary>
        ///     设置单元格的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public void Set(string name, string value)
        {
            Find(name).SetCellValue(value);
        }
        /// <summary>
        ///     设置单元格的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public void Set(string name, double value)
        {
            Find(name).SetCellValue(value);
        }
        /// <summary>
        ///     设置单元格的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public void Set(string name, bool value)
        {
            Find(name).SetCellValue(value);
        }
        /// <summary>
        ///     设置单元格的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public void Set(string name, DateTime value)
        {
            Find(name).SetCellValue(value);
        }
        private static void Set<T>(ICell cell, T value, Type type = null)
        {
            if (type == null)
            {
                type = value.GetType();
            }
            if (type == typeof (string))
            {
                cell.SetCellValue(value as string);
            }
            else if (type == typeof (DateTime))
            {
                cell.SetCellValue(Convert.ToDateTime(value));
            }
            else if (type == typeof (bool))
            {
                cell.SetCellValue(Convert.ToBoolean(value));
            }
            else if (type == typeof (double)
                     || type == typeof (float)
                     || type == typeof (int)
                     || type == typeof (uint)
                     || type == typeof (Int16)
                     || type == typeof (Int64)
                     || type == typeof (UInt16)
                     || type == typeof (UInt64))
            {
                cell.SetCellValue(Convert.ToDouble(value));
            }
            else
            {
                throw new Exception("cannot support type:" + type.FullName);
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
                throw new Exception("no member");
            }
            var row = FindRow(name);
            if (values.Count == 0)
            {
                if (willCopyRow)
                {
                    row.Sheet.RemoveRow(row);
                }
                else
                {
                    RemovePlaceHolder(name);
                }
                return;
            }
            if (willCopyRow)
            {
                for (var i = 0; i < values.Count - 1; i++)
                {
                    row.CopyRowTo(row.RowNum + 1);
                }
            }
            for (var i = 0; i < values.Count; i++)
            {
                var value = values[i];
                var nextRow = row.Sheet.GetRow(row.RowNum + 1);
                foreach (var propertyInfo in properties)
                {
                    var result = type.InvokeMember(propertyInfo.Name, BindingFlags.GetProperty, null, value, null);
                    var cell = Find(row, Combine(name, propertyInfo.Name));
                    if (!willCopyRow
                        && i != values.Count - 1)
                    {
                        var nextCell = nextRow.GetCell(cell.ColumnIndex);
                        if (nextCell == null)
                        {
                            nextCell = nextRow.CreateCell(cell.ColumnIndex);
                        }
                        nextCell.SetCellValue(cell.StringCellValue);
                    }
                    Set(cell, result);
                }
                foreach (var fieldInfo in fields)
                {
                    var result = type.InvokeMember(fieldInfo.Name, BindingFlags.GetField, null, value, null);
                    var cell = Find(row, Combine(name, fieldInfo.Name));
                    if (!willCopyRow
                        && i != values.Count - 1)
                    {
                        var nextCell = nextRow.GetCell(cell.ColumnIndex);
                        if (nextCell == null)
                        {
                            nextCell = nextRow.CreateCell(cell.ColumnIndex);
                        }
                        nextCell.SetCellValue(cell.StringCellValue);
                    }
                    Set(cell, result);
                }
                row = nextRow;
            }
        }
        private ICell Find(string name)
        {
            name = GetPlaceHolderName(name);
            for (var i = 0; i < _workbook.NumberOfSheets; i++)
            {
                var sheet = _workbook[i];
                for (var j = 0; j <= sheet.LastRowNum; j++)
                {
                    var row = sheet.GetRow(j);
                    if (row == null)
                    {
                        row = sheet.CreateRow(j);
                    }
                    for (var k = 0; k < row.LastCellNum; k++)
                    {
                        var cell = row.GetCell(k);
                        if (cell == null)
                        {
                            cell = row.CreateCell(k);
                        }
                        if (cell.CellType == CellType.String
                            && cell.StringCellValue.Trim() == name)
                        {
                            return cell;
                        }
                    }
                }
            }
            throw new Exception(name + " not found in excel");
        }
        private IRow FindRow(string name)
        {
            name = GetStartOfPlaceHolderName(name);
            for (var i = 0; i < _workbook.NumberOfSheets; i++)
            {
                var sheet = _workbook[i];
                for (var j = 0; j <= sheet.LastRowNum; j++)
                {
                    var row = sheet.GetRow(j);
                    if (row == null)
                    {
                        row = sheet.CreateRow(j);
                    }
                    for (var k = 0; k < row.LastCellNum; k++)
                    {
                        var cell = row.GetCell(k);
                        if (cell == null)
                        {
                            cell = row.CreateCell(k);
                        }
                        if (cell.CellType == CellType.String
                            && cell.StringCellValue.Trim().StartsWith(name))
                        {
                            return cell.Row;
                        }
                    }
                }
            }
            throw new Exception(name + " not found in excel");
        }
        private void RemovePlaceHolder(string name)
        {
            var row = FindRow(name);
            for (var k = 0; k < row.LastCellNum; k++)
            {
                var cell = row.GetCell(k);
                if (cell == null)
                {
                    cell = row.CreateCell(k);
                }
                if (cell.CellType == CellType.String
                    && cell.StringCellValue.Trim().StartsWith(GetStartOfPlaceHolderName(name)))
                {
                    row.RemoveCell(cell);
                }
            }
        }
        private static ICell Find(IRow row, string name)
        {
            name = GetPlaceHolderName(name);
            for (var k = 0; k < row.PhysicalNumberOfCells; k++)
            {
                var cell = row.Cells[k];
                if (cell.CellType == CellType.String
                    && cell.StringCellValue.Trim() == name)
                {
                    return cell;
                }
            }
            throw new Exception(name + " not found in row:" + row.RowNum);
        }
        private static string GetPlaceHolderName(string name)
        {
            return string.Format("{{{0}}}", name);
        }
        private static string Combine(string name, string memberName)
        {
            return name + "-" + memberName;
        }
        private static string GetStartOfPlaceHolderName(string name)
        {
            return string.Format("{{{0}-", name);
        }
        /// <summary>
        ///     远程下载Excel文件
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName"></param>
        public void Save(HttpResponse response, string fileName)
        {
            response.ContentType = "application/vnd.ms-excel";
            response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(fileName + ".xls", Encoding.UTF8));
            using (var stream = new MemoryStream())
            {
                _workbook.Write(stream);
                stream.Flush();
                stream.Position = 0;
                stream.WriteTo(response.OutputStream);
            }
        }
        /// <summary>
        ///     本地保存Excel文件
        /// </summary>
        /// <param name="file"></param>
        public void Save(string file)
        {
            using (var stream = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(stream);
            }
        }
    }
}