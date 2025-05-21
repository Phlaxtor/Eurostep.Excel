using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel;

public class DefaultFillStyle : FillStyle
{
    public DefaultFillStyle()
    {
        PatternType = PatternValues.None;
    }
}