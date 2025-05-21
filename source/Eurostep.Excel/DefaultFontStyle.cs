namespace Eurostep.Excel;

public class DefaultFontStyle : FontStyle
{
    public DefaultFontStyle()
    {
        Size = DefaultValue.FontSize;
        Color = new GeneralColor(DefaultValue.FontColor);
        Name = DefaultValue.FontName;
    }
}