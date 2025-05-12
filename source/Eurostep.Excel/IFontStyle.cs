using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel;

public interface IFontStyle
{
    bool? Bold { get; }
    GeneralColor? Color { get; }
    bool? Condense { get; }
    bool? Extend { get; }
    bool? Italic { get; }
    string? Name { get; }
    bool? Shadow { get; }
    double? Size { get; }
    bool? Strike { get; }
    UnderlineValues? Underline { get; }
    uint Value { get; }
    VerticalAlignmentRunValues? VerticalTextAlignment { get; }
}