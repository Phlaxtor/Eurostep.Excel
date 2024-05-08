using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    internal sealed class CellEnumerator : ElementEnumerator<ICell>
    {
        public CellEnumerator(Row row, ExcelContext context) : base(row.GetEnumerator(), context)
        {
        }

        protected override bool GetCurrent(out ICell? current)
        {
            current = default;
            if (Enumerator.Current is not Cell c) return false;
            current = new ExcelCell(c, Context);
            return true;
        }
    }
}