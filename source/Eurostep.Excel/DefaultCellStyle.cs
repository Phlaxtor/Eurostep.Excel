namespace Eurostep.Excel;

public class DefaultCellStyle : CellStyle
{
    public DefaultCellStyle()
    {
        Border = new DefaultBorderStyle();
        Fill = new DefaultFillStyle();
        Font = new DefaultFontStyle();
        NumberingFormat = new DefaultNumberingFormatStyle();
    }
}