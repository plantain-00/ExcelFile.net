using System;
using System.Web;

using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ExcelFile.net
{
    /// <summary>
    ///     <para>
    ///         ���ݣ�������Sheet()����Row()����Ԫ��Cell()���յĵ�Ԫ��Empty()���ϲ���Ԫ��Cell()
    ///     </para>
    ///     <para>
    ///         ��Ԫ����ʽ��Ĭ����ʽStyle������ʽNewStyle()��������ʽCell()������ʽRow()
    ///     </para>
    ///     <para>
    ///         ����ʽ���п�Sheet()
    ///     </para>
    ///     <para>
    ///         ����ʽ��Ĭ���и�DefaultRowHeight()�������и�Row()
    ///     </para>
    ///     <para>
    ///         ����������ļ�Save()��Զ������Save()
    ///     </para>
    /// </summary>
    public interface IExcelFile
    {
        /// <summary>
        ///     ���Ĭ����ʽ
        /// </summary>
        ExcelStyle Style { get; }

        /// <summary>
        ///     �½���ʽ
        /// </summary>
        /// <returns></returns>
        ExcelStyle NewStyle();

        /// <summary>
        ///     �½�������
        /// </summary>
        /// <param name="columnWidths"></param>
        /// <returns></returns>
        ExcelFile Sheet(params int[] columnWidths);

        /// <summary>
        ///     �½�������
        /// </summary>
        /// <param name="name"></param>
        /// <param name="widths"></param>
        /// <returns></returns>
        ExcelFile Sheet(string name, params int[] widths);

        /// <summary>
        ///     Ĭ���и�
        /// </summary>
        /// <param name="height"></param>
        /// <returns></returns>
        ExcelFile DefaultRowHeight(int height);

        /// <summary>
        ///     �½���
        /// </summary>
        /// <param name="rowStyle"></param>
        /// <returns></returns>
        ExcelFile Row(ExcelStyle rowStyle = null);

        /// <summary>
        ///     �½���
        /// </summary>
        /// <param name="height"></param>
        /// <param name="rowStyle"></param>
        /// <returns></returns>
        ExcelFile Row(short height, ExcelStyle rowStyle = null);

        /// <summary>
        ///     �½��յ�Ԫ��
        /// </summary>
        /// <param name="colspan"></param>
        /// <returns></returns>
        ExcelFile Empty(int colspan = 1);

        /// <summary>
        ///     �½���Ԫ��
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(string value, ExcelStyle cellStyle = null);

        /// <summary>
        ///     �½���Ԫ��
        /// </summary>
        /// <param name="value"></param>
        /// <param name="rowspan"></param>
        /// <param name="colspan"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(string value, int rowspan, int colspan, ExcelStyle cellStyle = null);

        /// <summary>
        ///     �½���Ԫ��
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(double value, ExcelStyle cellStyle = null);

        /// <summary>
        ///     �½���Ԫ��
        /// </summary>
        /// <param name="value"></param>
        /// <param name="rowspan"></param>
        /// <param name="colspan"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(double value, int rowspan, int colspan, ExcelStyle cellStyle = null);

        /// <summary>
        ///     �½���Ԫ��
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(bool value, ExcelStyle cellStyle = null);

        /// <summary>
        ///     �½���Ԫ��
        /// </summary>
        /// <param name="value"></param>
        /// <param name="rowspan"></param>
        /// <param name="colspan"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(bool value, int rowspan, int colspan, ExcelStyle cellStyle = null);

        /// <summary>
        ///     �½���Ԫ��
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(DateTime value, ExcelStyle cellStyle = null);

        /// <summary>
        ///     �½���Ԫ��
        /// </summary>
        /// <param name="value"></param>
        /// <param name="rowspan"></param>
        /// <param name="colspan"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        ExcelFile Cell(DateTime value, int rowspan, int colspan, ExcelStyle cellStyle = null);

        /// <summary>
        ///     Զ������Excel�ļ���MVC��return new EmptyResult();
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName">����չ��</param>
        void Save(HttpResponse response, string fileName);

        /// <summary>
        ///     ���ر���Excel�ļ�
        /// </summary>
        /// <param name="file">����չ��</param>
        void Save(string file);

#if !NET20 &&!NET30 &&!NET35
        /// <summary>
        ///     Զ������Excel�ļ���MVC��return new EmptyResult();
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName">����չ��</param>
        void Save(HttpResponseBase response, string fileName);
#endif
    }

    /// <summary>
    ///     <para>
    ///         ���ݣ�������Sheet()����Row()����Ԫ��Cell()���յĵ�Ԫ��Empty()���ϲ���Ԫ��Cell()
    ///     </para>
    ///     <para>
    ///         ��Ԫ����ʽ��Ĭ����ʽStyle������ʽNewStyle()��������ʽCell()������ʽRow()
    ///     </para>
    ///     <para>
    ///         ����ʽ���п�Sheet()
    ///     </para>
    ///     <para>
    ///         ����ʽ��Ĭ���и�DefaultRowHeight()�������и�Row()
    ///     </para>
    ///     <para>
    ///         ����������ļ�Save()��Զ������Save()
    ///     </para>
    /// </summary>
    public class ExcelFile : IExcelFile
    {
        /// <summary>
        ///     ��ǰ������
        /// </summary>
        public readonly IWorkbook Workbook;
        private readonly ICellStyle _cellStyle;
        private ICell _cell;
        private IRow _row;
        private ICellStyle _rowStyle;
        private ISheet _sheet;

        /// <summary>
        ///     ����Excel�ļ�����
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
        ///     ���Ĭ����ʽ
        /// </summary>
        public ExcelStyle Style => new ExcelStyle(_cellStyle, Workbook.CreateFont());

        /// <summary>
        ///     �½���ʽ
        /// </summary>
        /// <returns></returns>
        public ExcelStyle NewStyle()
        {
            return new ExcelStyle(Workbook.CreateCellStyle(), Workbook.CreateFont());
        }

        /// <summary>
        ///     �½�������
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
        ///     �½�������
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
        ///     Ĭ���и�
        /// </summary>
        /// <param name="height"></param>
        /// <returns></returns>
        public ExcelFile DefaultRowHeight(int height)
        {
            _sheet.DefaultRowHeightInPoints = height;
            return this;
        }

        /// <summary>
        ///     �½���
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
        ///     �½���
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
        ///     �½��յ�Ԫ��
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
        ///     �½���Ԫ��
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
        ///     �½���Ԫ��
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
        ///     �½���Ԫ��
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
        ///     �½���Ԫ��
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
        ///     �½���Ԫ��
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
        ///     �½���Ԫ��
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
        ///     �½���Ԫ��
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
        ///     �½���Ԫ��
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
        ///     Զ������Excel�ļ���MVC��return new EmptyResult();
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName">����չ��</param>
        public void Save(HttpResponse response, string fileName)
        {
            ExcelUtils.Save(Workbook, response, fileName);
        }

        /// <summary>
        ///     ���ر���Excel�ļ�
        /// </summary>
        /// <param name="file">����չ��</param>
        public void Save(string file)
        {
            ExcelUtils.Save(Workbook, file);
        }

#if !NET20 &&!NET30 &&!NET35
        /// <summary>
        ///     Զ������Excel�ļ���MVC��return new EmptyResult();
        /// </summary>
        /// <param name="response"></param>
        /// <param name="fileName">����չ��</param>
        public void Save(HttpResponseBase response, string fileName)
        {
            ExcelUtils.Save(Workbook, response, fileName);
        }
#endif
    }
}