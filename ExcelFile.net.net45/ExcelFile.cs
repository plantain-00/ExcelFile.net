using System;
using System.IO;
using System.Text;
using System.Web;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace ExcelFile.net
{
    /// <summary>
    ///     <para>
    ///         内容：工作表Sheet()、行Row()、单元格Cell()、空的单元格Empty()、合并单元格Cell()
    ///     </para>
    ///     <para>
    ///         单元格样式：默认样式Style、新样式NewStyle()、内联样式Cell()、行样式Row()
    ///     </para>
    ///     <para>
    ///         列样式：列宽Sheet()
    ///     </para>
    ///     <para>
    ///         行样式：默认行高DefaultRowHeight()、内联行高Row()
    ///     </para>
    ///     <para>
    ///         输出：本地文件Save()、远程下载Save()
    ///     </para>
    /// </summary>
    public class ExcelFile
    {
        private readonly ICellStyle _cellStyle;
        private readonly IWorkbook _workbook;
        private ICell _cell;
        private IRow _row;
        private ICellStyle _rowStyle;
        private ISheet _sheet;
        public ExcelFile(bool is2007OrMore = false)
        {
            _workbook = is2007OrMore ? new XSSFWorkbook() as IWorkbook : new HSSFWorkbook();
            _cellStyle = _workbook.CreateCellStyle();
            _cellStyle.Alignment = HorizontalAlignment.Center;
            _cellStyle.VerticalAlignment = VerticalAlignment.Top;
        }
        public ExcelStyle Style
        {
            get
            {
                return new ExcelStyle(_cellStyle, _workbook.CreateFont());
            }
        }
        public ExcelStyle NewStyle()
        {
            return new ExcelStyle(_workbook.CreateCellStyle(), _workbook.CreateFont());
        }
        public ExcelFile Sheet(params int[] widths)
        {
            _sheet = _workbook.CreateSheet();
            for (var i = 0; i < widths.Length; i++)
            {
                _sheet.SetColumnWidth(i, widths[i] * 256);
            }
            _row = null;
            _cell = null;
            return this;
        }
        public ExcelFile Sheet(string name, params int[] widths)
        {
            _sheet = _workbook.CreateSheet(name);
            for (var i = 0; i < widths.Length; i++)
            {
                _sheet.SetColumnWidth(i, widths[i] * 256);
            }
            _row = null;
            _cell = null;
            return this;
        }
        public ExcelFile DefaultRowHeight(int height)
        {
            _sheet.DefaultRowHeightInPoints = height;
            return this;
        }
        public ExcelFile Row(ExcelStyle rowStyle = null)
        {
            _row = _sheet.CreateRow(_row == null ? 0 : _row.RowNum + 1);
            _cell = null;
            _rowStyle = rowStyle == null ? null : rowStyle.Style;
            return this;
        }
        public ExcelFile Row(short height, ExcelStyle rowStyle = null)
        {
            Row(rowStyle);
            _row.Height = (short)(height * 20);
            return this;
        }
        public ExcelFile Empty(int colspan = 1)
        {
            for (var i = 0; i < colspan; i++)
            {
                Cell(string.Empty);
            }
            return this;
        }
        public ExcelFile Cell(string value, ExcelStyle cellStyle = null)
        {
            _cell = _row.CreateCell(_cell == null ? 0 : _cell.ColumnIndex + 1);
            _cell.SetCellValue(value);
            if (cellStyle != null)
            {
                _cell.CellStyle = cellStyle.Style;
            }
            else if (_rowStyle != null)
            {
                _cell.CellStyle = _rowStyle;
            }
            else
            {
                _cell.CellStyle = _cellStyle;
            }
            return this;
        }
        public ExcelFile Cell(string value, int rowspan, int colspan, ExcelStyle cellStyle = null)
        {
            Cell(value, cellStyle);
            Merge(_row.RowNum, _cell.ColumnIndex, rowspan, colspan);
            Empty(colspan - 1);
            return this;
        }
        public ExcelFile Cell(double value, ExcelStyle cellStyle = null)
        {
            _cell = _row.CreateCell(_cell == null ? 0 : _cell.ColumnIndex + 1);
            _cell.SetCellValue(value);
            if (cellStyle != null)
            {
                _cell.CellStyle = cellStyle.Style;
            }
            else if (_rowStyle != null)
            {
                _cell.CellStyle = _rowStyle;
            }
            else
            {
                _cell.CellStyle = _cellStyle;
            }
            return this;
        }
        public ExcelFile Cell(double value, int rowspan, int colspan, ExcelStyle cellStyle = null)
        {
            Cell(value, cellStyle);
            Merge(_row.RowNum, _cell.ColumnIndex, rowspan, colspan);
            Empty(colspan - 1);
            return this;
        }
        public ExcelFile Cell(bool value, ExcelStyle cellStyle = null)
        {
            _cell = _row.CreateCell(_cell == null ? 0 : _cell.ColumnIndex + 1);
            _cell.SetCellValue(value);
            if (cellStyle != null)
            {
                _cell.CellStyle = cellStyle.Style;
            }
            else if (_rowStyle != null)
            {
                _cell.CellStyle = _rowStyle;
            }
            else
            {
                _cell.CellStyle = _cellStyle;
            }
            return this;
        }
        public ExcelFile Cell(bool value, int rowspan, int colspan, ExcelStyle cellStyle = null)
        {
            Cell(value, cellStyle);
            Merge(_row.RowNum, _cell.ColumnIndex, rowspan, colspan);
            Empty(colspan - 1);
            return this;
        }
        public ExcelFile Cell(DateTime value, ExcelStyle cellStyle = null)
        {
            _cell = _row.CreateCell(_cell == null ? 0 : _cell.ColumnIndex + 1);
            _cell.SetCellValue(value);
            if (cellStyle != null)
            {
                _cell.CellStyle = cellStyle.Style;
            }
            else if (_rowStyle != null)
            {
                _cell.CellStyle = _rowStyle;
            }
            else
            {
                _cell.CellStyle = _cellStyle;
            }
            return this;
        }
        public ExcelFile Cell(DateTime value, int rowspan, int colspan, ExcelStyle cellStyle = null)
        {
            Cell(value, cellStyle);
            Merge(_row.RowNum, _cell.ColumnIndex, rowspan, colspan);
            Empty(colspan - 1);
            return this;
        }
        private void Merge(int row, int column, int rowspan, int colspan)
        {
            _sheet.AddMergedRegion(new CellRangeAddress(row, row + rowspan - 1, column, column + colspan - 1));
        }
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
        public void Save(string file)
        {
            using (var stream = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(stream);
            }
        }
    }
}