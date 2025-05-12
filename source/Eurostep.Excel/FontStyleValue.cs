using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel;

public readonly struct FontStyleValue : IFontStyle
{
    public FontStyleValue(uint value, string name, double? sz, bool? b, bool? i, UnderlineValues? u, GeneralColor? color, VerticalAlignmentRunValues? vertAlig, bool? strike, bool? condense, bool? extend, bool? shadow)
    {
        Bold = b;
        Color = color;
        Condense = condense;
        Extend = extend;
        Italic = i;
        Name = name ?? "Calibri";
        Shadow = shadow;
        Size = sz;
        Strike = strike;
        Underline = u;
        Value = value;
        VerticalTextAlignment = vertAlig;
    }

    public bool? Bold { get; }

    public GeneralColor? Color { get; }

    public bool? Condense { get; }

    public bool? Extend { get; }

    public bool? Italic { get; }

    public string? Name { get; }

    public bool? Shadow { get; }

    public double? Size { get; }

    public bool? Strike { get; }

    public UnderlineValues? Underline { get; }

    public uint Value { get; }

    public VerticalAlignmentRunValues? VerticalTextAlignment { get; }

    public static implicit operator uint(FontStyleValue value)
    {
        return value.Value;
    }
}