namespace Eurostep.Excel;

public class DefaultBorderStyle : BorderStyle
{
    public DefaultBorderStyle()
    {
        BottomBorder = new BorderPart();
        LeftBorder = new BorderPart();
        RightBorder = new BorderPart();
        TopBorder = new BorderPart();
    }
}