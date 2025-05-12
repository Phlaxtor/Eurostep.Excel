using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel;

public abstract class FormatStyle
{
    public Alignment? Alignment { get; init; }
    public IBorderStyle? Border { get; init; }
    public IFillStyle? Fill { get; init; }
    public IFontStyle? Font { get; init; }
    public uint? FormatId { get; init; }
    public INumberingFormatStyle? NumberingFormat { get; init; }
    public bool? PivotButton { get; init; }
    public Protection? Protection { get; init; }
    public bool? QuotePrefix { get; init; }
}