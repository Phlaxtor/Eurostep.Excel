using System.Diagnostics.CodeAnalysis;

namespace Eurostep.Excel;

public readonly struct HeaderId : IEquatable<HeaderId>, IComparable<HeaderId>
{
    public static readonly HeaderId Empty = new HeaderId();

    public HeaderId()
    {
        Column = ColumnId.Empty;
        Id = string.Empty;
        Name = string.Empty;
    }

    public HeaderId(string name, ColumnId column)
    {
        Column = column;
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Id = name.GetToUpperWithoutWhiteSpace();
    }

    public ColumnId Column { get; }

    public string Id { get; }

    public string Name { get; }

    public static implicit operator string(HeaderId c)
    {
        return c.Name;
    }

    public static implicit operator uint(HeaderId c)
    {
        return c.Column.No;
    }

    public static bool operator !=(HeaderId l, HeaderId r)
    {
        return l.Column != r.Column;
    }

    public static bool operator <(HeaderId l, HeaderId r)
    {
        return l.Column < r.Column;
    }

    public static bool operator <=(HeaderId l, HeaderId r)
    {
        return l.Column <= r.Column;
    }

    public static bool operator ==(HeaderId l, HeaderId r)
    {
        return l.Column == r.Column;
    }

    public static bool operator >(HeaderId l, HeaderId r)
    {
        return l.Column > r.Column;
    }

    public static bool operator >=(HeaderId l, HeaderId r)
    {
        return l.Column >= r.Column;
    }

    public int CompareTo(HeaderId other)
    {
        return Column.CompareTo(other.Column);
    }

    public override bool Equals([NotNullWhen(true)] object? obj)
    {
        if (obj is HeaderId other)
        {
            return Equals(other);
        }
        return false;
    }

    public bool Equals(HeaderId other)
    {
        return Column.Equals(other.Column);
    }

    public override int GetHashCode()
    {
        return Column.GetHashCode();
    }

    public override string ToString()
    {
        return $"{Name} ({Column})";
    }
}