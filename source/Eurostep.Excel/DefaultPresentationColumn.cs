namespace Eurostep.Excel
{
    public sealed class DefaultPresentationColumn : IPresentationColumn
    {
        public DefaultPresentationColumn(string displayName, double width, CellStyleValue? styleIndex = default, CellStyleValue? columnStyle = default)
        {
            ColumnStyle = columnStyle;
            DisplayName = displayName;
            HeaderStyle = styleIndex;
            Width = width;
        }

        public CellStyleValue? ColumnStyle { get; }
        public string DisplayName { get; }
        public CellStyleValue? HeaderStyle { get; }
        public double Width { get; }
    }
}