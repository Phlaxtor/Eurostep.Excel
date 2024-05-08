namespace Eurostep.Excel
{
    internal sealed class RowReader : ElementReader<IRow, ICell>, IRowReader
    {
        public RowReader(IRow row, ExcelContext context) : base(row, context)
        {
        }
    }
}