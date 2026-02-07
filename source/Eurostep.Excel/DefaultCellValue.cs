using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    public sealed class DefaultCellValue : ICellValue
    {
        public DefaultCellValue(string? value, CellStyleValue? style, CellValues dataType)
        {
            Value = value;
            Style = style;
            DataType = dataType;
        }

        public CellValues DataType { get; }
        public CellStyleValue? Style { get; }
        public string? Value { get; }

        public static ICellValue[] Get(params string?[] values)
        {
            ICellValue[] result = new ICellValue[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                result[i] = new DefaultCellValue(values[i], null, CellValues.String);
            }
            return result;
        }
    }
}