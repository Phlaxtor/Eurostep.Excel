namespace Eurostep.Excel;

public interface INumberingFormatStyle
{
    string? FormatCode { get; }
    uint? NumberFormatId { get; }
}