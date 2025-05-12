namespace Eurostep.Excel;

public readonly struct FormatStyleValue
{
    public FormatStyleValue()
    {
    }

    public uint Value { get; }

    public static implicit operator uint(FormatStyleValue value)
    {
        return value.Value;
    }
}