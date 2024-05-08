namespace Eurostep.Excel
{
    internal sealed class ExcelWriterData
    {
        private IPresentationColumn[]? _headers;

        public ExcelWriterData(string name, uint sheetNo, ColumnId columnStart, uint rowStart)
        {
            var sheetName = GetSheetName(name);
            ColumnStart = columnStart;
            Name = name;
            RowEnd = rowStart - 1;
            RowStart = rowStart;
            SheetName = sheetName;
            SheetNo = sheetNo;
            Start = new CellRef(columnStart, rowStart, sheetName);
        }

        public ColumnId ColumnEnd => GetColumnEnd();
        public ColumnId ColumnStart { get; }
        public IPresentationColumn[] Headers => GetHeaders();
        public string Name { get; }
        public uint RowEnd { get; private set; }
        public uint RowStart { get; }
        public string SheetName { get; }
        public uint SheetNo { get; }
        public CellRef Start { get; }
        public uint TableEnd { get; private set; }
        public uint TableId { get; private set; }
        public uint TableStart { get; private set; }

        public uint EndTable()
        {
            TableEnd = RowEnd;
            return TableEnd;
        }

        public CellArea GetArea() => new CellArea(ColumnStart, RowStart, ColumnEnd, RowEnd);

        public CellRef GetCurrentCell(uint columnOffset = 0)
        {
            if (columnOffset == 0) return new CellRef(ColumnStart, RowEnd, SheetName);
            return new CellRef(ColumnStart + columnOffset, RowEnd, SheetName);
        }

        public CellRef GetEnd() => new CellRef(ColumnEnd, RowEnd, SheetName);

        public CellArea GetTableArea() => new CellArea(ColumnStart, TableStart, ColumnEnd, TableEnd);

        public uint IncreaseRowNo()
        {
            RowEnd++;
            return RowEnd;
        }

        public void SetHeaders(IPresentationColumn[] value)
        {
            if (_headers != null) throw new ApplicationException($"Headers is already set for {SheetName}");
            _headers = value;
        }

        public uint StartTable(uint tableId)
        {
            TableId = tableId;
            TableStart = RowEnd;
            return TableStart;
        }

        private ColumnId GetColumnEnd()
        {
            if (_headers == null) return ColumnStart;
            if (_headers.Length == 0) return ColumnStart;
            return ColumnStart.No + (uint)_headers.Length - 1;
        }

        private IPresentationColumn[] GetHeaders()
        {
            if (_headers != null) return _headers;
            return Array.Empty<IPresentationColumn>();
        }

        private string GetSheetName(string name)
        {
            if (name.Length < 32) return name;
            return name.Remove(32);
        }
    }
}