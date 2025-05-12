using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel;

public abstract class FillStyle : IFillStyle
{
    public GeneralColor? BackgroundColor { get; init; }

    public GeneralColor? ForegroundColor { get; init; }

    public double GradientBottom { get; init; }

    public double GradientDegree { get; init; }

    public double GradientLeft { get; init; }

    public double GradientRight { get; init; }

    public double GradientTop { get; init; }

    public GradientValues? GradientType { get; init; }

    public PatternValues? PatternType { get; init; }
}