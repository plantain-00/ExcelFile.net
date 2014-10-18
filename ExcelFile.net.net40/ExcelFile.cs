using System;
using System.Web;

using NPOI.SS.UserModel;
using NPOI.SS.Util;

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
        /// <summary>
        ///     当前工作簿
        /// </summary>
        public readonly IWorkbook Workbook;
        private readonly ICellStyle _cellStyle;
        private ICell _cell;
        private IRow _row;
        private ICellStyle _rowStyle;
        private ISheet _sheet;

        /// <summary>
        ///     构造Excel文件对象
        /// </summary>
        /// <param name="is2007OrLater"></param>
        public ExcelFile(bool is2007OrLater = false)
        {
            Workbook = ExcelUtils.New(is2007OrLater);
            _cellStyle = Workbook.CreateCellStyle();
            _cellStyle.Alignment = HorizontalAlignment.Center;
            _cellStyle.VerticalAlignment = VerticalAlignment.Top;
        }

        /// <summary>
        ///     获得默认样式
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                return new ExcelStyle(_cellStyle, Workbook.CreateFont());
            }
        }

        /// <summary>
        ///     新建样式
        /// </summary>
        /// <returns></returns>
        public ExcelStyle NewStyle()
        {
            return new ExcelStyle(Workbook.CreateCellStyle(), Workbook.CreateFont());
        }

        /// <summary>
        ///     新建工作表
        /// </summary>
        /// <param name="columnWidths"></param>
        /// <returns></returns>
        public ExcelFile Sheet(params int[] columnWidths)
        {
            _sheet = Workbook.CreateSheet();
            for (var i = 0; i < columnWidths.Length; i++)
            {
                _sheet.SetColumnWidth(i, columnWidths[i] * 256);
            }
            _row = null;
            _cell = null;
            return this;
        }

        /// <summary>
        ///     新建工作表
        /// </summary>
        /// <param name="name"></param>
        /// <param name="widths"></param>
        /// <returns></returns>
        public ExcelFile Sheet(string name, params int[] widths)
        {
            _sheet = Workbook.CreateSheet(name);
            for (var i = 0; i < widths.Length; i++)
            {
                _sheet.SetColumnWidth(i, widths[i] * 256);
            }
            _row = null;
            _cell = null;
            return this;
        }

        /// <summary>
        ///     默认行高
        /// </summary>
        /// <param name="height"></param>
        /// <returns></returns>
        public ExcelFile DefaultRowHeight(int height)
        {
            _sheet.DefaultRowHeightInPoints = height;
            return this;
        }

        /// <summary>
        ///     新建行
        /// </summary>
        /// <param name="rowStyle"></param>
        /// <returns></returns>
        public ExcelFile Row(ExcelStyle rowStyle = null)
        {
            _row = _sheet.CreateRow(_row == null ? 0 : _row.RowNum + 1);
            _cell = null;
            _rowStyle = rowStyle == null ? null : rowStyle.Style;
            return this;
        }

        /// <summary>
        ///     新建行
        /// </summary>
        /// <param name="height"></param>
        /// <param name="rowStyle"></param>
        /// <returns></returns>
        public ExcelFile Row(short height, ExcelStyle rowStyle = null)
        {
            Row(rowStyle);
            _row.Height = (short) (height * 20);
            return this;
        }

        /// <summary>
        ///     新建空单元格
        /// </summary>
        /// <param name="colspan"></param>
        /// <returns></returns>
        public ExcelFile Empty(int colspan = 1)
        {
            for (var i = 0; i < colspan; i++)
            {
                Cell(string.Empty);
            }
            return this;
        }

        /// <summary>
        ///     新建单元格
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
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

        /// <summary>
        ///     新建单元格
        /// </summary>
        /// <param name="value"></param>
        /// <param name="rowspan"></param>
        /// <param name="colspan"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public ExcelFile Cell(string value, int rowspan, int colspan, ExcelStyle cellStyle = null)
        {
            Cell(value, cellStyle);
            Merge(_row.RowNum, _cell.ColumnIndex, rowspan, colspan);
            Empty(colspan - 1);
            return this;
        }

        /// <summary>
        ///     新建单元格
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
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

        /// <summary>
        ///     新建单元格
        /// </summary>
        /// <param name="value"></param>
        /// <param name="rowspan"></param>
        /// <param name="colspan"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public ExcelFile Cell(double value, int rowspan, int colspan, ExcelStyle cellStyle = null)
        {
            Cell(value, cellStyle);
            Merge(_row.RowNum, _cell.ColumnIndex, rowspan, colspan);
            Empty(colspan - 1);
            return this;
        }

        /// <summary>
        ///     新建单元格
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
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

        /// <summary>
        ///     新建单元格
        /// </summary>
        /// <param name="value"></param>
        /// <param name="rowspan"></param>
        /// <param name="colspan"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public ExcelFile Cell(bool value, int rowspan, int colspan, ExcelStyle cellStyle = null)
        {
            Cell(value, cellStyle);
            Merge(_row.RowNum, _cell.ColumnIndex, rowspan, colspan);
            Empty(colspan - 1);
            return this;
        }

        /// <summary>
        ///     新建单元格
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
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

        /// <summary>
        ///     新建单元格
        /// </summary>
        /// <param name="value"></param>
        /// <param name="rowspan"></param>
        /// <param name="colspan"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
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

        /// <summary>
        ///     远程下载Excel文件，MVC中return new EmptyResult();
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName">带扩展名</param>
        public void Save(HttpResponse response, string fileName)
        {
            ExcelUtils.Save(Workbook, response, fileName);
        }

        /// <summary>
        ///     本地保存Excel文件
        /// </summary>
        /// <param name="file">带扩展名</param>
        public void Save(string file)
        {
            ExcelUtils.Save(Workbook, file);
        }

#if !NET20 &&!NET30 &&!NET35
        /// <summary>
        ///     远程下载Excel文件，MVC中return new EmptyResult();
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName">带扩展名</param>
        public void Save(HttpResponseBase response, string fileName)
        {
            ExcelUtils.Save(Workbook, response, fileName);
        }
#endif
    }
}