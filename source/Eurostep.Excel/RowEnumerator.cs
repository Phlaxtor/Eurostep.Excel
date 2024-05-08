using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    internal sealed class RowEnumerator : ElementEnumerator<IRow>
    {
        public RowEnumerator(SheetData sheetData, ExcelContext context) : base(sheetData.GetEnumerator(), context)
        {
        }

        protected override bool GetCurrent(out IRow? current)
        {
            current = default;
            if (Enumerator.Current is not Row c) return false;
            current = new ExcelRow(c, Context);
            return true;
        }
    }
}