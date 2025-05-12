namespace Eurostep.Excel;

public readonly struct NumberingFormatStyleValue : INumberingFormatStyle
{
    public NumberingFormatStyleValue(uint value, uint? numberFormatId, string? formatCode)
    {
        FormatCode = formatCode;
        NumberFormatId = numberFormatId;
        Value = value;
    }

    public string? FormatCode { get; }

    public uint? NumberFormatId { get; }

    public uint Value { get; }

    public static implicit operator uint(NumberingFormatStyleValue value)
    {
        return value.Value;
    }
}