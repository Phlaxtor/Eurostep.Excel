using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;

namespace Eurostep.Excel
{
    public readonly struct CellRef
    {
        private const string ColumnName = "Column";
        private const string RowName = "Row";
        private static readonly Regex _regex = new Regex("(?<Column>[a-zA-Z]*)(?<Row>[0-9]*)");

        public CellRef(string column, uint rowId, string? sheetId = null)
        {
            Column = column;
            RowId = rowId;
            SheetId = sheetId;
        }

        public CellRef(uint column, uint rowId, string? sheetId = null)
        {
            Column = column;
            RowId = rowId;
            SheetId = sheetId;
        }

        public CellRef(ColumnId column, uint rowId, string? sheetId = null)
        {
            Column = column;
            RowId = rowId;
            SheetId = sheetId;
        }

        public ColumnId Column { get; }

        public uint RowId { get; }

        public string? SheetId { get; }

        public static implicit operator string(CellRef c)
        {
            return c.ToString();
        }

        public static bool operator !=(CellRef l, CellRef r)
        {
            return l.Column != r.Column || l.RowId != r.RowId || l.SheetId != r.SheetId;
        }

        public static bool operator <(CellRef l, CellRef r)
        {
            if (l.RowId < r.RowId) return true;
            if (l.Column < r.Column) return true;
            return false;
        }

        public static bool operator <=(CellRef l, CellRef r)
        {
            if (l.Column == r.Column && l.RowId == r.RowId) return true;
            return l < r;
        }

        public static bool operator ==(CellRef l, CellRef r)
        {
            return l.Column == r.Column && l.RowId == r.RowId && l.SheetId == r.SheetId;
        }

        public static bool operator >(CellRef l, CellRef r)
        {
            if (l.RowId > r.RowId) return true;
            if (l.Column > r.Column) return true;
            return false;
        }

        public static bool operator >=(CellRef l, CellRef r)
        {
            if (l.Column == r.Column && l.RowId == r.RowId) return true;
            return l > r;
        }

        public static CellRef Parse(string s)
        {
            Match match = _regex.Match(s);
            string column = match.Groups[ColumnName].Value;
            uint row = uint.Parse(match.Groups[RowName].Value);
            return new CellRef(column, row);
        }

        public static bool TryParse(string? s, [NotNullWhen(true)] out CellRef value)
        {
            value = default;
            if (string.IsNullOrEmpty(s)) return false;
            try
            {
                value = Parse(s);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public override bool Equals([NotNullWhen(true)] object? obj)
        {
            if (obj is not CellRef other) return false;
            if (other.Column != Column) return false;
            if (other.RowId != RowId) return false;
            if (other.SheetId != SheetId) return false;
            return true;
        }

        public CellRef GetForColumn(ColumnId column)
        {
            return new CellRef(column, RowId, SheetId);
        }

        public CellRef GetForRow(uint rowId)
        {
            return new CellRef(Column, rowId, SheetId);
        }

        public CellRef GetForSheet(string sheetId)
        {
            return new CellRef(Column, RowId, sheetId);
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        public override string ToString()
        {
            if (string.IsNullOrEmpty(SheetId)) return $"{Column}{RowId}";
            return $"'{SheetId}'!{Column}{RowId}";
        }
    }
}