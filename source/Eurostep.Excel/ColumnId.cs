using System.Diagnostics.CodeAnalysis;

namespace Eurostep.Excel
{
    public readonly struct ColumnId : IEquatable<ColumnId>, IComparable<ColumnId>
    {
        public static readonly ColumnId Empty = new ColumnId();
        public static readonly ColumnId MaxValue = new ColumnId(16384U);
        public static readonly ColumnId MinValue = new ColumnId(1U);

        public ColumnId()
        {
            ColumnName = ColumnName.None;
            Index = -1;
            Name = string.Empty;
            No = 0;
        }

        public ColumnId(string name)
        {
            No = GetColumnNo(name);
            Name = name;
            Index = GetColumnIndex(No);
            ColumnName = (ColumnName)No;
        }

        public ColumnId(uint no)
        {
            No = no;
            Name = GetColumnName(no);
            Index = GetColumnIndex(no);
            ColumnName = (ColumnName)no;
        }

        public ColumnId(ColumnName columnName)
        {
            No = (uint)columnName;
            Name = GetColumnName(No);
            Index = GetColumnIndex(No);
            ColumnName = columnName;
        }

        public ColumnId(int index)
        {
            No = GetColumnNo(index);
            Name = GetColumnName(index);
            Index = index;
            ColumnName = (ColumnName)No;
        }

        public ColumnName ColumnName { get; }

        public int Index { get; }

        public string Name { get; }

        public uint No { get; }

        public static implicit operator ColumnId(uint c)
        {
            return new ColumnId(c);
        }

        public static implicit operator ColumnId(string c)
        {
            return new ColumnId(c);
        }

        public static implicit operator string(ColumnId c)
        {
            return c.Name;
        }

        public static implicit operator uint(ColumnId c)
        {
            return c.No;
        }

        public static ColumnId operator -(ColumnId l, int r)
        {
            return new ColumnId(l.No - 1U);
        }

        public static ColumnId operator --(ColumnId c)
        {
            return new ColumnId(c.No - 1U);
        }

        public static bool operator !=(ColumnId l, ColumnId r)
        {
            return l.No != r.No;
        }

        public static ColumnId operator +(ColumnId l, int r)
        {
            return new ColumnId(l.No + 1U);
        }

        public static ColumnId operator ++(ColumnId c)
        {
            return new ColumnId(c.No + 1U);
        }

        public static bool operator <(ColumnId l, ColumnId r)
        {
            return l.No < r.No;
        }

        public static bool operator <=(ColumnId l, ColumnId r)
        {
            return l.No <= r.No;
        }

        public static bool operator ==(ColumnId l, ColumnId r)
        {
            return l.No == r.No;
        }

        public static bool operator >(ColumnId l, ColumnId r)
        {
            return l.No > r.No;
        }

        public static bool operator >=(ColumnId l, ColumnId r)
        {
            return l.No >= r.No;
        }

        public int CompareTo(ColumnId other)
        {
            return No.CompareTo(other.No);
        }

        public override bool Equals([NotNullWhen(true)] object? obj)
        {
            if (obj is ColumnId other)
            {
                return Equals(other);
            }
            return false;
        }

        public bool Equals(ColumnId other)
        {
            return No == other.No;
        }

        public CellRef GetCellRef(uint rowId, string? sheetId = null)
        {
            return new CellRef(this, rowId, sheetId);
        }

        public override int GetHashCode()
        {
            return Index;
        }

        public override string ToString()
        {
            return Name;
        }

        private static int GetColumnIndex(uint no)
        {
            return ((int)no) - 1;
        }

        private static string GetColumnName(uint no)
        {
            string result = string.Empty;
            int r = (int)no;
            while (r > 0)
            {
                r--;
                int i = (r % 26);
                r = (r / 26);
                char c = (char)(i + 65);
                result = $"{c}{result}";
            }
            return result;
        }

        private static string GetColumnName(int index)
        {
            string result = string.Empty;
            int r = index;
            while (r >= 0)
            {
                r--;
                int i = (r % 26);
                r = (r / 26);
                char c = (char)(i + 65);
                result = $"{c}{result}";
            }
            return result;
        }

        private static uint GetColumnNo(int index)
        {
            return ((uint)index) + 1;
        }

        private static uint GetColumnNo(string column)
        {
            uint result = 0;
            int position = 0;
            for (int i = column.Length - 1; i >= 0; i--)
            {
                int index = char.ToUpper(column[i]) - 64;
                result += (uint)(Math.Pow(26, position) * index);
                position++;
            }
            return result;
        }
    }
}