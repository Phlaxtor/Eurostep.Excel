namespace Eurostep.Excel;

public interface IBorderStyle
{
    BorderPart? BottomBorder { get; }
    BorderPart? LeftBorder { get; }
    BorderPart? RightBorder { get; }
    BorderPart? TopBorder { get; }
}