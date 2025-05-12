namespace Eurostep.Excel;

public abstract class NumberingFormatStyle : INumberingFormatStyle
{
    public string? FormatCode { get; init; }

    public uint? NumberFormatId { get; init; }
}