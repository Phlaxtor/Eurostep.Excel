using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel;

public abstract class FontStyle : IFontStyle
{
    public bool? Bold { get; init; }

    public GeneralColor? Color { get; init; }

    public bool? Condense { get; init; }

    public bool? Extend { get; init; }

    public bool? Italic { get; init; }

    public string? Name { get; init; }

    public bool? Shadow { get; init; }

    public double? Size { get; init; }

    public bool? Strike { get; init; }

    public UnderlineValues? Underline { get; init; }

    public uint Value { get; init; }

    public VerticalAlignmentRunValues? VerticalTextAlignment { get; init; }
}