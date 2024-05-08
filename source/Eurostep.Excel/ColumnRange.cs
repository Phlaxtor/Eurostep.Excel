using System.Diagnostics.CodeAnalysis;

namespace Eurostep.Excel
{
    public readonly struct ColumnRange
    {
        public const uint RowMax = 1048576;
        public const uint RowMin = 1;

        public ColumnRange(string column, uint startRow = RowMin, uint endRow = RowMax, string? sheetId = null)
        {
            Column = column;
            StartRow = startRow;
            EndRow = endRow;
            SheetId = sheetId;
        }

        public ColumnRange(uint column, uint startRow = RowMin, uint endRow = RowMax, string? sheetId = null)
        {
            Column = column;
            StartRow = startRow;
            EndRow = endRow;
            SheetId = sheetId;
        }

        public ColumnRange(ColumnId column, uint startRow = RowMin, uint endRow = RowMax, string? sheetId = null)
        {
            Column = column;
            StartRow = startRow;
            EndRow = endRow;
            SheetId = sheetId;
        }

        public ColumnId Column { get; }
        public uint EndRow { get; }
        public string? SheetId { get; }
        public uint StartRow { get; }

        public static implicit operator string(ColumnRange c) => c.ToString();

        public override bool Equals([NotNullWhen(true)] object? obj)
        {
            if (obj is not ColumnRange other) return false;
            if (other.Column != Column) return false;
            if (other.StartRow != StartRow) return false;
            if (other.EndRow != EndRow) return false;
            if (other.SheetId != SheetId) return false;
            return true;
        }

        public string Get(RefType column, RefType startRow, RefType endRow)
        {
            if (string.IsNullOrEmpty(SheetId)) return $"{ColumnRef(column)}{StartRowRef(startRow)}:{ColumnRef(column)}{EndRowRef(endRow)}";
            return $"'{SheetId}'!{ColumnRef(column)}{StartRowRef(startRow)}:{ColumnRef(column)}{EndRowRef(endRow)}";
        }

        public string GetAbsolute() => Get(RefType.Absolute, RefType.Absolute, RefType.Absolute);

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        public string GetRelative() => Get(RefType.Relative, RefType.Relative, RefType.Relative);

        public override string ToString() => GetAbsolute();

        private string ColumnRef(RefType type)
        {
            return type switch
            {
                RefType.Absolute => '$' + Column.ToString(),
                RefType.Relative => Column.ToString(),
                _ => throw new ApplicationException($"Not supported {nameof(RefType)} '{type}'")
            };
        }

        private string EndRowRef(RefType type)
        {
            return type switch
            {
                RefType.Absolute => '$' + EndRow.ToString(),
                RefType.Relative => EndRow.ToString(),
                _ => throw new ApplicationException($"Not supported {nameof(RefType)} '{type}'")
            };
        }

        private string StartRowRef(RefType type)
        {
            return type switch
            {
                RefType.Absolute => '$' + StartRow.ToString(),
                RefType.Relative => StartRow.ToString(),
                _ => throw new ApplicationException($"Not supported {nameof(RefType)} '{type}'")
            };
        }
    }
}