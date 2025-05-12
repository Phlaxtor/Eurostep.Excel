using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace Eurostep.Excel
{
    internal sealed class ExcelRow : ExcelElement, IRow
    {
        private readonly Row _row;
        private int? _length;

        public ExcelRow(Row row, ExcelContext context) : base(context)
        {
            _row = row;
            RowIndex = row.RowIndex?.Value ?? throw new ArgumentNullException(nameof(RowIndex));
        }

        public int Length => GetLength();
        public uint RowIndex { get; }
        public ICell this[string column] => GetCellByColumnId(column) ?? throw new ApplicationException($"Can not find cell with column id '{column}'");
        public ICell this[uint index] => GetCellByColumnIndex(index) ?? throw new ApplicationException($"Can not find cell at index '{index}'");

        public ICell? GetCellByColumnId(string column)
        {
            ICell? result = GetCellByReference(column);
            if (result != null) return result;
            return default;
        }

        public ICell? GetCellByColumnIndex(uint index)
        {
            string column = GetColumnName(index);
            ICell? result = GetCellByReference(column);
            if (result != null) return result;
            return default;
        }

        public IEnumerator<ICell> GetEnumerator()
        {
            return new CellEnumerator(_row, Context);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return new CellEnumerator(_row, Context);
        }

        public IRowReader GetReader()
        {
            return new RowReader(this, Context);
        }

        public ITableRow GetTableRow(IHeaderLookup lookup)
        {
            return new TableRow(this, lookup, Context);
        }

        public string?[] GetValues()
        {
            string?[] result = new string[Length];
            using IEnumerator<ICell> enumerator = GetEnumerator();
            while (enumerator.MoveNext())
            {
                ICell c = enumerator.Current;
                result[c.Index] = c.Value;
            }
            return result;
        }

        protected override int GetIndex()
        {
            return ((int)RowIndex) - 1;
        }

        protected override bool GetIsEmpty()
        {
            using IEnumerator<ICell> enumerator = GetEnumerator();
            while (enumerator.MoveNext())
            {
                if (enumerator.Current.IsEmpty == false) return false;
            }
            return true;
        }

        private ICell? GetCellByReference(string column)
        {
            string reference = $"{column}{RowIndex}";
            foreach (DocumentFormat.OpenXml.OpenXmlElement element in _row)
            {
                if (element is not Cell cell) continue;
                if (cell.CellReference is null) continue;
                if (cell.CellReference.Value != reference) continue;
                return new ExcelCell(cell, Context);
            }
            return default;
        }

        private int GetLength()
        {
            if (_length.HasValue) return _length.Value;
            if (_row.HasChildren == false) return 0;
            if (_row.LastChild is null) return 0;
            if (_row.LastChild is not Cell l) throw new ApplicationException("LastChild on Row is not a Cell");
            string r = l.CellReference?.Value ?? throw new ArgumentNullException(nameof(Cell.CellReference));
            (string Column, uint Row) reference = ParseReference(r);
            int index = GetColumnIndex(reference.Column);
            _length = index + 1;
            return _length.Value;
        }
    }
}