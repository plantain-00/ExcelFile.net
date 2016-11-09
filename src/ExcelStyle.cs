using NPOI.SS.UserModel;

namespace ExcelFile.net
{
    /// <summary>
    ///     <para>
    ///         Background: Background
    ///     </para>
    ///     <para>
    ///         Border and the color of border: Border, BorderTop, BorderBottom, BorderLeft, BorderRight
    ///     </para>
    ///     <para>
    ///         Align: Align, VerticalAlign
    ///     </para>
    ///     <para>
    ///         Word: WrapText, Italic, Underline, FontSize, Font, Color, Bold
    ///     </para>
    /// </summary>
    public interface IExcelStyle
    {
        /// <summary>
        ///     Get style of cell
        /// </summary>
        ICellStyle Style { get; }

        /// <summary>
        ///     Background
        /// </summary>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle Background(short HSSFColor);

        /// <summary>
        ///     BorderTop
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle BorderTop(BorderStyle borderStyle, short HSSFColor = -1);

        /// <summary>
        ///     BorderBottom
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle BorderBottom(BorderStyle borderStyle, short HSSFColor = -1);

        /// <summary>
        ///     BorderLeft
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle BorderLeft(BorderStyle borderStyle, short HSSFColor = -1);

        /// <summary>
        ///     BorderRight
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle BorderRight(BorderStyle borderStyle, short HSSFColor = -1);

        /// <summary>
        ///     Border
        /// </summary>
        /// <param name="borderStyle"></param>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle Border(BorderStyle borderStyle, short HSSFColor = -1);

        /// <summary>
        ///     Align
        /// </summary>
        /// <param name="alignment"></param>
        /// <returns></returns>
        ExcelStyle Align(HorizontalAlignment alignment);

        /// <summary>
        ///     VerticalAlign
        /// </summary>
        /// <param name="alignment"></param>
        /// <returns></returns>
        ExcelStyle VerticalAlign(VerticalAlignment alignment);

        /// <summary>
        ///     WrapText
        /// </summary>
        /// <param name="wrapText"></param>
        /// <returns></returns>
        ExcelStyle WrapText(bool wrapText);

        /// <summary>
        ///     Color
        /// </summary>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        ExcelStyle Color(short HSSFColor);

        /// <summary>
        ///     Italic
        /// </summary>
        /// <returns></returns>
        ExcelStyle Italic();

        /// <summary>
        ///     Underline
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        ExcelStyle Underline(FontUnderlineType type);

        /// <summary>
        ///     FontSize
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        ExcelStyle FontSize(double size);

        /// <summary>
        ///     Font
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        ExcelStyle Font(string name);

        /// <summary>
        ///     Bold
        /// </summary>
        /// <returns></returns>
        ExcelStyle Bold();
    }

    /// <summary>
    ///     <para>
    ///         Background: Background
    ///     </para>
    ///     <para>
    ///         Border and the color of border: Border, BorderTop, BorderBottom, BorderLeft, BorderRight
    ///     </para>
    ///     <para>
    ///         Align: Align, VerticalAlign
    ///     </para>
    ///     <para>
    ///         Word: WrapText, Italic, Underline, FontSize, Font, Color, Bold
    ///     </para>
    /// </summary>
    public class ExcelStyle : IExcelStyle
    {
        private readonly IFont _font;
        /// <summary>
        ///     Construct an ExcelStyle object
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
        ///     Construct an ExcelStyle object
        /// </summary>
        public ICellStyle Style { get; }
        /// <summary>
        ///     Background
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
        ///     BorderTop
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
        ///     BorderBottom
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
        ///     BorderLeft
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
        ///     BorderRight
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
        ///     Border
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
        ///     Align
        /// </summary>
        /// <param name="alignment"></param>
        /// <returns></returns>
        public ExcelStyle Align(HorizontalAlignment alignment)
        {
            Style.Alignment = alignment;
            return this;
        }
        /// <summary>
        ///     VerticalAlign
        /// </summary>
        /// <param name="alignment"></param>
        /// <returns></returns>
        public ExcelStyle VerticalAlign(VerticalAlignment alignment)
        {
            Style.VerticalAlignment = alignment;
            return this;
        }
        /// <summary>
        ///     WrapText
        /// </summary>
        /// <param name="wrapText"></param>
        /// <returns></returns>
        public ExcelStyle WrapText(bool wrapText)
        {
            Style.WrapText = wrapText;
            return this;
        }
        /// <summary>
        ///     Color
        /// </summary>
        /// <param name="HSSFColor"></param>
        /// <returns></returns>
        public ExcelStyle Color(short HSSFColor)
        {
            _font.Color = HSSFColor;
            return this;
        }
        /// <summary>
        ///     Italic
        /// </summary>
        /// <returns></returns>
        public ExcelStyle Italic()
        {
            _font.IsItalic = true;
            return this;
        }
        /// <summary>
        ///     Underline
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public ExcelStyle Underline(FontUnderlineType type)
        {
            _font.Underline = type;
            return this;
        }
        /// <summary>
        ///     FontSize
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public ExcelStyle FontSize(double size)
        {
            _font.FontHeight = size;
            return this;
        }
        /// <summary>
        ///     Font
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public ExcelStyle Font(string name)
        {
            _font.FontName = name;
            return this;
        }
        /// <summary>
        ///     Bold
        /// </summary>
        /// <returns></returns>
        public ExcelStyle Bold()
        {
            _font.Boldweight = (short) FontBoldWeight.Bold;
            return this;
        }
    }
}