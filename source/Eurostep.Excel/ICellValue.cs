using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    public interface ICellValue
    {
        CellValues DataType { get; }
        CellStyle? Style { get; }
        string? Value { get; }
    }
}