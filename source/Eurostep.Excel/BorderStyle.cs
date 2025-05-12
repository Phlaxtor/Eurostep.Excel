namespace Eurostep.Excel;

public abstract class BorderStyle : IBorderStyle
{
    public BorderPart? BottomBorder { get; init; }

    public BorderPart? LeftBorder { get; init; }

    public BorderPart? RightBorder { get; init; }

    public BorderPart? TopBorder { get; init; }
}