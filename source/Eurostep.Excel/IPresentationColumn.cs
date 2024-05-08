namespace Eurostep.Excel
{
    public interface IPresentationColumn
    {
        CellStyle? ColumnStyle { get; }
        string DisplayName { get; }
        CellStyle? HeaderStyle { get; }
        int Width { get; }
    }
}