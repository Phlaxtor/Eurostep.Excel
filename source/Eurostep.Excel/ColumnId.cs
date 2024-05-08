using System.Diagnostics.CodeAnalysis;

namespace Eurostep.Excel
{
    public readonly struct ColumnId
    {
        public static readonly ColumnId MaxValue = new ColumnId(16384);
        public static readonly ColumnId MinValue = new ColumnId(1);

        public ColumnId(string name)
        {
            Name = name;
            No = GetColumnNo(name);
            Index = GetColumnIndex(No);
        }

        public ColumnId(uint no)
        {
            No = no;
            Name = GetColumnName(no);
            Index = GetColumnIndex(no);
        }

        public int Index { get; }
        public string Name { get; }
        public uint No { get; }

        public static implicit operator ColumnId(uint c) => new ColumnId(c);

        public static implicit operator ColumnId(string c) => new ColumnId(c);

        public static implicit operator string(ColumnId c) => c.Name;

        public static implicit operator uint(ColumnId c) => c.No;

        public static ColumnId operator -(ColumnId l, int r) => new ColumnId(l.No - 1);

        public static ColumnId operator --(ColumnId c) => new ColumnId(c.No - 1);

        public static bool operator !=(ColumnId l, ColumnId r) => l.No != r.No;

        public static ColumnId operator +(ColumnId l, int r) => new ColumnId(l.No + 1);

        public static ColumnId operator ++(ColumnId c) => new ColumnId(c.No + 1);

        public static bool operator <(ColumnId l, ColumnId r) => l.No < r.No;

        public static bool operator <=(ColumnId l, ColumnId r) => l.No <= r.No;

        public static bool operator ==(ColumnId l, ColumnId r) => l.No == r.No;

        public static bool operator >(ColumnId l, ColumnId r) => l.No > r.No;

        public static bool operator >=(ColumnId l, ColumnId r) => l.No >= r.No;

        public override bool Equals([NotNullWhen(true)] object? obj)
        {
            if (obj is not ColumnId other) return false;
            if (other.No != No) return false;
            return true;
        }

        public CellRef GetCellRef(uint rowId, string? sheetId = null) => new CellRef(this, rowId, sheetId);

        public override int GetHashCode()
        {
            return unchecked((int)No);
        }

        public override string ToString()
        {
            return Name;
        }

        private static int GetColumnIndex(uint column) => ((int)column) - 1;

        private static string GetColumnName(uint index)
        {
            var result = string.Empty;
            int r = (int)index;
            while (r > 0)
            {
                int i = (r % 26);
                r = (r / 26);
                char c = (char)(i + 64);
                result = $"{c}{result}";
            }
            return result;
        }

        private static uint GetColumnNo(string column)
        {
            uint result = 0;
            int position = 0;
            for (int i = column.Length - 1; i >= 0; i--)
            {
                var index = char.ToUpper(column[i]) - 64;
                result += (uint)(Math.Pow(26, position) * index);
                position++;
            }
            return result;
        }
    }
}