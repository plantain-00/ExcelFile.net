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
    public class ExcelStyle
    {
        private readonly IFont _font;
        public ExcelStyle(ICellStyle style, IFont font)
        {
            Style = style;
            _font = font;
            Style.SetFont(_font);
        }
        public ICellStyle Style { get; private set; }
        public ExcelStyle Background(short HSSFColor)
        {
            Style.FillPattern = FillPattern.SolidForeground;
            Style.FillForegroundColor = HSSFColor;
            return this;
        }
        public ExcelStyle BorderTop(BorderStyle borderStyle, short HSSFColor = -1)
        {
            Style.BorderTop = borderStyle;
            if (HSSFColor != -1)
            {
                Style.TopBorderColor = HSSFColor;
            }
            return this;
        }
        public ExcelStyle BorderBottom(BorderStyle borderStyle, short HSSFColor = -1)
        {
            Style.BorderBottom = borderStyle;
            if (HSSFColor != -1)
            {
                Style.BottomBorderColor = HSSFColor;
            }
            return this;
        }
        public ExcelStyle BorderLeft(BorderStyle borderStyle, short HSSFColor = -1)
        {
            Style.BorderLeft = borderStyle;
            if (HSSFColor != -1)
            {
                Style.LeftBorderColor = HSSFColor;
            }
            return this;
        }
        public ExcelStyle BorderRight(BorderStyle borderStyle, short HSSFColor = -1)
        {
            Style.BorderRight = borderStyle;
            if (HSSFColor != -1)
            {
                Style.RightBorderColor = HSSFColor;
            }
            return this;
        }
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
        public ExcelStyle Align(HorizontalAlignment alignment)
        {
            Style.Alignment = alignment;
            return this;
        }
        public ExcelStyle VerticalAlign(VerticalAlignment alignment)
        {
            Style.VerticalAlignment = alignment;
            return this;
        }
        public ExcelStyle WrapText(bool wrapText)
        {
            Style.WrapText = wrapText;
            return this;
        }
        public ExcelStyle Color(short HSSFColor)
        {
            _font.Color = HSSFColor;
            return this;
        }
        public ExcelStyle Italic()
        {
            _font.IsItalic = true;
            return this;
        }
        public ExcelStyle Underline(FontUnderlineType type)
        {
            _font.Underline = type;
            return this;
        }
        public ExcelStyle FontSize(double size)
        {
            _font.FontHeight = size;
            return this;
        }
        public ExcelStyle Font(string name)
        {
            _font.FontName = name;
            return this;
        }
        public ExcelStyle Bold()
        {
            _font.Boldweight = (short)FontBoldWeight.Bold;
            return this;
        }
    }
}