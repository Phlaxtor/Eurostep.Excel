using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    public sealed class RowBuilder
    {
        private readonly ISheetWriter _excel;
        private readonly List<ICellValue> _row = new List<ICellValue>();

        internal RowBuilder(ISheetWriter excel)
        {
            _excel = excel;
        }

        public ISheetWriter Build()
        {
            _excel.AddRow(_row.ToArray());
            return _excel;
        }

        public RowBuilder New(string? value, CellStyle? style = null, CellValues dataType = CellValues.String)
        {
            _row.Add(new DefaultCellValue(value, style, dataType));
            return this;
        }
    }
}