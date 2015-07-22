using NPOI.SS.UserModel;

namespace ExcelFile.net
{
    /// <summary>
    ///     <para>
    ///         背景色：Background
    ///     </para>
    ///     <para>
    ///         边框及边框颜色：Border、BorderTop、BorderBottom、BorderLeft、BorderRight
    ///     </para>
    ///     <para>
    ///         对齐：Align、VerticalAlign
    ///     </para>
    ///     <para>
    ///         文字：WrapText、Italic、Underline、FontSize、Font、Color、Bold
    ///     </para>
    /// </summary>
    public interface IExcelStyle
    {
        /// <summary>
        ///     构造的Excel样式对象
        /// </summary>
        ICellStyle Style { get; }

        /// <summary>
        ///     背景色
        /// </summary>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle Background(short HSSFColor);

        /// <summary>
        ///     上边框
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle BorderTop(BorderStyle borderStyle, short HSSFColor = -1);

        /// <summary>
        ///     下边框
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle BorderBottom(BorderStyle borderStyle, short HSSFColor = -1);

        /// <summary>
        ///     左边框
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle BorderLeft(BorderStyle borderStyle, short HSSFColor = -1);

        /// <summary>
        ///     右边框
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle BorderRight(BorderStyle borderStyle, short HSSFColor = -1);

        /// <summary>
        ///     边框
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle Border(BorderStyle borderStyle, short HSSFColor = -1);

        /// <summary>
        ///     水平对齐
        /// </summary>
        /// <param name="alignment"></param>
        /// <returns></returns>
        ExcelStyle Align(HorizontalAlignment alignment);

        /// <summary>
        ///     垂直对齐
        /// </summary>
        /// <param name="alignment"></param>
        /// <returns></returns>
        ExcelStyle VerticalAlign(VerticalAlignment alignment);

        /// <summary>
        ///     文本自动换行
        /// </summary>
        /// <param name="wrapText"></param>
        /// <returns></returns>
        ExcelStyle WrapText(bool wrapText);

        /// <summary>
        ///     前景色
        /// </summary>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle Color(short HSSFColor);

        /// <summary>
        ///     斜体
        /// </summary>
        /// <returns></returns>
        ExcelStyle Italic();

        /// <summary>
        ///     下划线
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        ExcelStyle Underline(FontUnderlineType type);

        /// <summary>
        ///     字体尺寸
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        ExcelStyle FontSize(double size);

        /// <summary>
        ///     字体
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        ExcelStyle Font(string name);

        /// <summary>
        ///     加粗
        /// </summary>
        /// <returns></returns>
        ExcelStyle Bold();
    }

    /// <summary>
    ///     <para>
    ///         背景色：Background
    ///     </para>
    ///     <para>
    ///         边框及边框颜色：Border、BorderTop、BorderBottom、BorderLeft、BorderRight
    ///     </para>
    ///     <para>
    ///         对齐：Align、VerticalAlign
    ///     </para>
    ///     <para>
    ///         文字：WrapText、Italic、Underline、FontSize、Font、Color、Bold
    ///     </para>
    /// </summary>
    public class ExcelStyle : IExcelStyle
    {
        private readonly IFont _font;
        /// <summary>
        ///     构造Excel样式对象
        /// </summary>
        /// <param name="style"></param>
        /// <param name="font"></param>
        public ExcelStyle(ICellStyle style, IFont font)
        {
            Style = style;
            _font = font;
            Style.SetFont(_font);
        }
        /// <summary>
        ///     构造的Excel样式对象
        /// </summary>
        public ICellStyle Style { get; }
        /// <summary>
        ///     背景色
        /// </summary>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        public ExcelStyle Background(short HSSFColor)
        {
            Style.FillPattern = FillPattern.SolidForeground;
            Style.FillForegroundColor = HSSFColor;
            return this;
        }
        /// <summary>
        ///     上边框
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        public ExcelStyle BorderTop(BorderStyle borderStyle, short HSSFColor = -1)
        {
            Style.BorderTop = borderStyle;
            if (HSSFColor != -1)
            {
                Style.TopBorderColor = HSSFColor;
            }
            return this;
        }
        /// <summary>
        ///     下边框
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        public ExcelStyle BorderBottom(BorderStyle borderStyle, short HSSFColor = -1)
        {
            Style.BorderBottom = borderStyle;
            if (HSSFColor != -1)
            {
                Style.BottomBorderColor = HSSFColor;
            }
            return this;
        }
        /// <summary>
        ///     左边框
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        public ExcelStyle BorderLeft(BorderStyle borderStyle, short HSSFColor = -1)
        {
            Style.BorderLeft = borderStyle;
            if (HSSFColor != -1)
            {
                Style.LeftBorderColor = HSSFColor;
            }
            return this;
        }
        /// <summary>
        ///     右边框
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        public ExcelStyle BorderRight(BorderStyle borderStyle, short HSSFColor = -1)
        {
            Style.BorderRight = borderStyle;
            if (HSSFColor != -1)
            {
                Style.RightBorderColor = HSSFColor;
            }
            return this;
        }
        /// <summary>
        ///     边框
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        public ExcelStyle Border(BorderStyle borderStyle, short HSSFColor = -1)
        {
            Style.BorderTop = borderStyle;
            Style.BorderBottom = borderStyle;
            Style.BorderLeft = borderStyle;
            Style.BorderRight = borderStyle;
            if (HSSFColor != -1)
            {
                Style.TopBorderColor = HSSFColor;
                Style.BottomBorderColor = HSSFColor;
                Style.LeftBorderColor = HSSFColor;
                Style.RightBorderColor = HSSFColor;
            }
            return this;
        }
        /// <summary>
        ///     水平对齐
        /// </summary>
        /// <param name="alignment"></param>
        /// <returns></returns>
        public ExcelStyle Align(HorizontalAlignment alignment)
        {
            Style.Alignment = alignment;
            return this;
        }
        /// <summary>
        ///     垂直对齐
        /// </summary>
        /// <param name="alignment"></param>
        /// <returns></returns>
        public ExcelStyle VerticalAlign(VerticalAlignment alignment)
        {
            Style.VerticalAlignment = alignment;
            return this;
        }
        /// <summary>
        ///     文本自动换行
        /// </summary>
        /// <param name="wrapText"></param>
        /// <returns></returns>
        public ExcelStyle WrapText(bool wrapText)
        {
            Style.WrapText = wrapText;
            return this;
        }
        /// <summary>
        ///     前景色
        /// </summary>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        public ExcelStyle Color(short HSSFColor)
        {
            _font.Color = HSSFColor;
            return this;
        }
        /// <summary>
        ///     斜体
        /// </summary>
        /// <returns></returns>
        public ExcelStyle Italic()
        {
            _font.IsItalic = true;
            return this;
        }
        /// <summary>
        ///     下划线
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public ExcelStyle Underline(FontUnderlineType type)
        {
            _font.Underline = type;
            return this;
        }
        /// <summary>
        ///     字体尺寸
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public ExcelStyle FontSize(double size)
        {
            _font.FontHeight = size;
            return this;
        }
        /// <summary>
        ///     字体
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public ExcelStyle Font(string name)
        {
            _font.FontName = name;
            return this;
        }
        /// <summary>
        ///     加粗
        /// </summary>
        /// <returns></returns>
        public ExcelStyle Bold()
        {
            _font.Boldweight = (short) FontBoldWeight.Bold;
            return this;
        }
    }
}