using System;
using System.Web;

using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ExcelFile.net
{
    /// <summary>
    ///     <para>
    ///         Content: Sheet(), Row(), Cell(), Empty()
    ///     </para>
    ///     <para>
    ///         Style of cell: Style, NewStyle(), Cell()¡¢Row()
    ///     </para>
    ///     <para>
    ///         Style of column: Sheet()
    ///     </para>
    ///     <para>
    ///         Style of row: DefaultRowHeight(), Row()
    ///     </para>
    ///     <para>
    ///         I/O: Save()
    ///     </para>
    /// </summary>
    public interface IExcelFile
    {
        /// <summary>
        ///     Default style
        /// </summary>
        ExcelStyle Style { get; }

        /// <summary>
        ///     New style
        /// </summary>
        /// <returns></returns>
        ExcelStyle NewStyle();

        /// <summary>
        ///     New worksheet
        /// </summary>
        /// <param name="columnWidths"></param>
        /// <returns></returns>
        ExcelFile Sheet(params int[] columnWidths);

        /// <summary>
        ///     New worksheet
        /// </summary>
        /// <param name="name"></param>
        /// <param name="widths"></param>
        /// <returns></returns>
        ExcelFile Sheet(string name, params int[] widths);

        /// <summary>
        ///     Default row height
        /// </summary>
        /// <param name="height"></param>
        /// <returns></returns>
        ExcelFile DefaultRowHeight(int height);

        /// <summary>
        ///     New row
        /// </summary>
        /// <param name="rowStyle"></param>
        /// <returns></returns>
        ExcelFile Row(ExcelStyle rowStyle = null);

        /// <summary>
        ///     New row
        /// </summary>
        /// <param name="height"></param>
        /// <param name="rowStyle"></param>
        /// <returns></returns>
        ExcelFile Row(short height, ExcelStyle rowStyle = null);

        /// <summary>
        ///     New empty cell
        /// </summary>
        /// <param name="colspan"></param>
        /// <returns></returns>
        ExcelFile Empty(int colspan = 1);

        /// <summary>
        ///     New cell
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(string value, ExcelStyle cellStyle = null);

        /// <summary>
        ///     New cell
        /// </summary>
        /// <param name="value"></param>
        /// <param name="rowspan"></param>
        /// <param name="colspan"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(string value, int rowspan, int colspan, ExcelStyle cellStyle = null);

        /// <summary>
        ///     New cell
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(double value, ExcelStyle cellStyle = null);

        /// <summary>
        ///     New cell
        /// </summary>
        /// <param name="value"></param>
        /// <param name="rowspan"></param>
        /// <param name="colspan"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(double value, int rowspan, int colspan, ExcelStyle cellStyle = null);

        /// <summary>
        ///     New cell
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(bool value, ExcelStyle cellStyle = null);

        /// <summary>
        ///     New cell
        /// </summary>
        /// <param name="value"></param>
        /// <param name="rowspan"></param>
        /// <param name="colspan"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(bool value, int rowspan, int colspan, ExcelStyle cellStyle = null);

        /// <summary>
        ///     New cell
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(DateTime value, ExcelStyle cellStyle = null);

        /// <summary>
        ///     New cell
        /// </summary>
        /// <param name="value"></param>
        /// <param name="rowspan"></param>
        /// <param name="colspan"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(DateTime value, int rowspan, int colspan, ExcelStyle cellStyle = null);

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
#endif
    }

    /// <summary>
    ///     <para>
    ///         Content: Sheet(), Row(), Cell(), Empty()
    ///     </para>
    ///     <para>
    ///         Style of cell: Style, NewStyle(), Cell()¡¢Row()
    ///     </para>
    ///     <para>
    ///         Style of column: Sheet()
    ///     </para>
    ///     <para>
    ///         Style of row: DefaultRowHeight(), Row()
    ///     </para>
    ///     <para>
    ///         I/O: Save()
    ///     </para>
    /// </summary>
    public class ExcelFile : IExcelFile
    {
        /// <summary>
        ///     Current workbook
        /// </summary>
        public readonly IWorkbook Workbook;
        private readonly ICellStyle _cellStyle;
        private ICell _cell;
        private IRow _row;
        private ICellStyle _rowStyle;
        private ISheet _sheet;

        /// <summary>
        ///     Construct an new ExcelFile object
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
        ///     Default style
        /// </summary>
        public ExcelStyle Style => new ExcelStyle(_cellStyle, Workbook.CreateFont());

        /// <summary>
        ///     New style
        /// </summary>
        /// <returns></returns>
        public ExcelStyle NewStyle()
        {
            return new ExcelStyle(Workbook.CreateCellStyle(), Workbook.CreateFont());
        }

        /// <summary>
        ///     New worksheet
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
        ///     New worksheet
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
        ///     Default row height
        /// </summary>
        /// <param name="height"></param>
        /// <returns></returns>
        public ExcelFile DefaultRowHeight(int height)
        {
            _sheet.DefaultRowHeightInPoints = height;
            return this;
        }

        /// <summary>
        ///     New row
        /// </summary>
        /// <param name="rowStyle"></param>
        /// <returns></returns>
        public ExcelFile Row(ExcelStyle rowStyle = null)
        {
            _row = _sheet.CreateRow(_row?.RowNum + 1 ?? 0);
            _cell = null;
            _rowStyle = rowStyle?.Style;
            return this;
        }

        /// <summary>
        ///     New row
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
        ///     New empty cell
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
        ///     New cell
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public ExcelFile Cell(string value, ExcelStyle cellStyle = null)
        {
            _cell = _row.CreateCell(_cell?.ColumnIndex + 1 ?? 0);
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
        ///     New cell
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
        ///     New cell
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public ExcelFile Cell(double value, ExcelStyle cellStyle = null)
        {
            _cell = _row.CreateCell(_cell?.ColumnIndex + 1 ?? 0);
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
        ///     New cell
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
        ///     New cell
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public ExcelFile Cell(bool value, ExcelStyle cellStyle = null)
        {
            _cell = _row.CreateCell(_cell?.ColumnIndex + 1 ?? 0);
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
        ///     New cell
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
        ///     New cell
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public ExcelFile Cell(DateTime value, ExcelStyle cellStyle = null)
        {
            _cell = _row.CreateCell(_cell?.ColumnIndex + 1 ?? 0);
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
        ///     New cell
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
#endif
    }
}