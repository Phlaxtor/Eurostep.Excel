using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel;

public readonly struct FillStyleValue : IFillStyle
{
    public FillStyleValue(uint value, PatternValues? patternType, GeneralColor? foregroundColor, GeneralColor? backgroundColor, GradientValues? gradientType, double gradientDegree, double gradientTop, double gradientBottom, double gradientRight, double gradientLeft)
    {
        BackgroundColor = backgroundColor;
        ForegroundColor = foregroundColor;
        GradientBottom = gradientBottom;
        GradientDegree = gradientDegree;
        GradientLeft = gradientLeft;
        GradientRight = gradientRight;
        GradientTop = gradientTop;
        GradientType = gradientType;
        PatternType = patternType;
        Value = value;
    }

    public GeneralColor? BackgroundColor { get; }

    public GeneralColor? ForegroundColor { get; }

    public double GradientBottom { get; }

    public double GradientDegree { get; }

    public double GradientLeft { get; }

    public double GradientRight { get; }

    public double GradientTop { get; }

    public GradientValues? GradientType { get; }

    public PatternValues? PatternType { get; }

    public uint Value { get; }

    public static implicit operator uint(FillStyleValue value)
    {
        return value.Value;
    }
}