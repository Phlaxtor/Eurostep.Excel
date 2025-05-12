using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel;

public interface IFillStyle
{
    GeneralColor? BackgroundColor { get; }
    GeneralColor? ForegroundColor { get; }
    double GradientBottom { get; }
    double GradientDegree { get; }
    double GradientLeft { get; }
    double GradientRight { get; }
    double GradientTop { get; }
    GradientValues? GradientType { get; }
    PatternValues? PatternType { get; }
}