using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    internal sealed class SheetEnumerator : ElementEnumerator<ISheet>
    {
        public SheetEnumerator(Sheets sheets, ExcelContext context) : base(sheets.GetEnumerator(), context)
        {
        }

        protected override bool GetCurrent(out ISheet? current)
        {
            current = default;
            if (Enumerator.Current is not Sheet c) return false;
            current = new ExcelSheet(c, Context, Position);
            return true;
        }
    }
}