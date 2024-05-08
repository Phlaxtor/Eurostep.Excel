using System.Collections;

namespace Eurostep.Excel
{
    internal sealed class TableRow : ExcelElement, ITableRow
    {
        private readonly IHeaderLookup _lookup;
        private readonly IRow _row;
        private readonly ICell?[] _values;
        private bool _hasValues = false;

        public TableRow(IRow row, IHeaderLookup lookup, ExcelContext context) : base(context)
        {
            _row = row;
            _lookup = lookup;
            _values = new ICell[Length];
        }

        public int Length => _row.Length;

        public string? this[string header] => GetText(header);

        public bool? GetBoolean(string header)
        {
            ICell? value = GetCell(header);
            if (value == null) return default;
            return value.GetBoolean();
        }

        public ICell? GetCell(string header)
        {
            int index = _lookup.GetIndex(header);
            return GetCell(index);
        }

        public ICell? GetCell(int index)
        {
            if (index < 0) return null;
            if (_values.Length <= index) return null;
            SetValues();
            return _values[index];
        }

        public DateTime? GetDateTime(string header)
        {
            ICell? value = GetCell(header);
            if (value == null) return default;
            return value.GetDateTime();
        }

        public IEnumerator<ICell> GetEnumerator() => _row.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => _row.GetEnumerator();

        public double? GetNumber(string header)
        {
            ICell? value = GetCell(header);
            if (value == null) return default;
            return value.GetNumber();
        }

        public string? GetText(string header)
        {
            ICell? value = GetCell(header);
            if (value == null) return default;
            return value.GetText();
        }

        protected override int GetIndex() => _row.Index;

        protected override bool GetIsEmpty() => _row.IsEmpty;

        private void SetValues()
        {
            if (_hasValues) return;
            using var enumerator = GetEnumerator();
            while (enumerator.MoveNext())
            {
                var c = enumerator.Current;
                _values[c.Index] = c;
            }
            _hasValues = true;
        }
    }
}