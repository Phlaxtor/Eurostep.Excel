using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel;

public interface ICellStyle
{
    Alignment? Alignment { get; }
    IBorderStyle? Border { get; }
    IFillStyle? Fill { get; }
    IFontStyle? Font { get; }
    uint? FormatId { get; }
    INumberingFormatStyle? NumberingFormat { get; }
    bool? PivotButton { get; }
    Protection? Protection { get; }
    bool? QuotePrefix { get; }
}