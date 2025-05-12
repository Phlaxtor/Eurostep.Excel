namespace Eurostep.Excel;

internal sealed class CellStyleCache : ExcelStyleCache<ICellStyle, CellStyleValue>
{
    public CellStyleCache(ExcelWriter writer) : base(writer)
    {
    }

    protected override CellStyleValue Create(ICellStyle input)
    {
        ArgumentNullException.ThrowIfNull(input, nameof(input));
        NumberingFormatStyleValue? numberingFormat = Writer.NumberingFormats.GetOrDefault(input.NumberingFormat);
        FontStyleValue font = Writer.FontStyles.GetOrDefault(input.Font);
        BorderStyleValue border = Writer.BorderStyles.GetOrDefault(input.Border);
        FillStyleValue fill = Writer.FillStyles.GetOrDefault(input.Fill);
        return Writer.NewCellStyle(input.Name, numberingFormat, input.FormatId, input.Alignment, font, border, fill, input.Protection, input.PivotButton, input.QuotePrefix);
    }
}