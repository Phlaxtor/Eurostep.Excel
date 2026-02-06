using DocumentFormat.OpenXml.Spreadsheet;

namespace FMA.Digimat.Toolbox.Excel.Model
{
    public sealed class CellText
    {
        public CellText(string value, ExcelFontStyle style)
        {
            Value = value;
            Style = style;
        }

        public ExcelFontStyle Style { get; }

        public string Value { get; }

        public static CellText New(string value, ExcelFontStyle style)
        {
            return new CellText(value, style);
        }

        public static CellText New<T>(string value) where T : ExcelFontStyle, new()
        {
            return new CellText(value, new T());
        }
    }

    public sealed class DefaultExcelBorderStyle : ExcelBorderStyle
    {
    }

    public sealed class DefaultExcelCellStyle : ExcelCellStyle
    {
        public DefaultExcelCellStyle()
        {
            BorderStyle = new DefaultExcelBorderStyle();
            FillStyle = new DefaultExcelFillStyle();
            FontStyle = new DefaultExcelFontStyle();
            NumberingFormatStyle = new DefaultExcelNumberingFormatStyle();
        }
    }

    public sealed class DefaultExcelFillStyle : ExcelFillStyle
    {
    }

    public sealed class DefaultExcelFontStyle : ExcelFontStyle
    {
    }

    public sealed class DefaultExcelNumberingFormatStyle : ExcelNumberingFormatStyle
    {
    }

    public abstract class ExcelBorderStyle
    {
        public string BorderColor { get; init; } = DefaultValue.BorderColor;
        public BorderStyleValues BorderStyle { get; init; } = BorderStyleValues.None;
    }

    public abstract class ExcelCellStyle
    {
        public Alignment? Alignment { get; init; }
        public ExcelBorderStyle? BorderStyle { get; init; }
        public ExcelFillStyle? FillStyle { get; init; }
        public ExcelFontStyle? FontStyle { get; init; }
        public ExcelNumberingFormatStyle? NumberingFormatStyle { get; set; }
        public bool? PivotButton { get; init; }
        public Protection? Protection { get; init; }
        public bool? QuotePrefix { get; init; }
    }

    public abstract class ExcelFillStyle
    {
        public string? BackgroundColor { get; init; }
        public string? ForegroundColor { get; init; }
        public PatternValues PatternType { get; init; } = PatternValues.None;
    }

    public abstract class ExcelFontStyle
    {
        public string FontColor { get; init; } = DefaultValue.FontColor;
        public string FontName { get; init; } = DefaultValue.FontName;
        public double? FontSize { get; init; } = DefaultValue.FontSize;
        public bool? HasShadow { get; init; }
        public bool? IsBold { get; init; }
        public bool? IsCondense { get; init; }
        public bool? IsExtend { get; init; }
        public bool? IsItalic { get; init; }
        public bool? IsStrike { get; init; }
        public UnderlineValues? UnderlineType { get; init; }
        public VerticalAlignmentRunValues? VerticalAlignment { get; init; }
    }

    public abstract class ExcelNumberingFormatStyle
    {
        public string? FormatCode { get; init; }
        public uint? NumberFormatId { get; init; }
    }
}