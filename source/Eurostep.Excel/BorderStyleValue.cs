namespace Eurostep.Excel;

public readonly struct BorderStyleValue : IBorderStyle
{
    public BorderStyleValue(uint value, BorderPart? top, BorderPart? right, BorderPart? bottom, BorderPart? left)
    {
        BottomBorder = bottom;
        LeftBorder = left;
        RightBorder = right;
        TopBorder = top;
        Value = value;
    }

    public BorderPart? BottomBorder { get; }

    public BorderPart? LeftBorder { get; }

    public BorderPart? RightBorder { get; }

    public BorderPart? TopBorder { get; }

    public uint Value { get; }

    public static implicit operator uint(BorderStyleValue value)
    {
        return value.Value;
    }
}