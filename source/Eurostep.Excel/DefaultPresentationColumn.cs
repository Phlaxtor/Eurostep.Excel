namespace Eurostep.Excel
{
    public sealed class DefaultPresentationColumn : IPresentationColumn
    {
        public DefaultPresentationColumn(string displayName, int width, CellStyle? styleIndex = default, CellStyle? columnStyle = default)
        {
            ColumnStyle = columnStyle;
            DisplayName = displayName;
            HeaderStyle = styleIndex;
            Width = width;
        }

        public CellStyle? ColumnStyle { get; }
        public string DisplayName { get; }
        public CellStyle? HeaderStyle { get; }
        public int Width { get; }
    }
}