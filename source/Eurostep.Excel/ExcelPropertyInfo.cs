using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace Eurostep.Excel;

public sealed class ExcelPropertyInfo : IEquatable<ExcelPropertyInfo>, IComparable<ExcelPropertyInfo>
{
    private readonly PropertyInfo _value;
    private ColumnId? _columnId;

    public ExcelPropertyInfo(PropertyInfo value, int index)
    {
        Index = index;
        _value = value ?? throw new ArgumentNullException(nameof(value));
        foreach (ExcelAttribute attribute in value.GetAttributes<ExcelAttribute>())
        {
            if (attribute is ExcelIgnoreAttribute)
            {
                Ignore = true;
                continue;
            }
            if (attribute is ExcelHeaderAttribute header)
            {
                HeaderName = header.Name;
                HeaderDescription = header.Description;
                continue;
            }
            if (attribute is ExcelColumnAttribute column)
            {
                ProvidedColumnName = column.Column;
                Width = column.Width;
                continue;
            }
            if (attribute is ExcelRequiredAttribute)
            {
                Required = true;
                continue;
            }
            if (attribute is ExcelStylesheetAttribute style)
            {
                switch (style.StyleType)
                {
                    case ExcelStyleType.Cell:
                        CellStyle = style.GetStylesheetDefinition();
                        break;

                    case ExcelStyleType.Header:
                        HeaderStyle = style.GetStylesheetDefinition();
                        break;

                    case ExcelStyleType.HeaderDescription:
                        HeaderDescriptionStyle = style.GetStylesheetDefinition();
                        break;
                }
                continue;
            }
        }

        Id = Name.GetToUpperWithoutWhiteSpace();
    }

    public ExcelStylesheetDefinition? CellStyle { get; }

    public ColumnId Column => GetColumnId();

    [MemberNotNullWhen(true, nameof(CellStyle))]
    public bool HasCellStyle { get; }

    [MemberNotNullWhen(true, nameof(HeaderDescription))]
    public bool HasHeaderDescription { get; }

    [MemberNotNullWhen(true, nameof(HeaderDescriptionStyle))]
    public bool HasHeaderDescriptionStyle { get; }

    [MemberNotNullWhen(true, nameof(HeaderName))]
    public bool HasHeaderName { get; }

    [MemberNotNullWhen(true, nameof(HeaderStyle))]
    public bool HasHeaderStyle { get; }

    public string? HeaderDescription { get; }

    public ExcelStylesheetDefinition? HeaderDescriptionStyle { get; }

    public string? HeaderName { get; }

    public ExcelStylesheetDefinition? HeaderStyle { get; }

    public string Id { get; }

    public bool Ignore { get; }

    public int Index { get; }

    public string Name => HeaderName ?? PropertyName;

    public string PropertyName => _value.Name;

    public ColumnName ProvidedColumnName { get; }

    public bool Required { get; }

    public double Width { get; }

    public int CompareTo(ExcelPropertyInfo? other)
    {
        if (other is null)
        {
            throw new ArgumentNullException(nameof(other));
        }
        return Column.CompareTo(other.Column);
    }

    public bool Equals(ExcelPropertyInfo? other)
    {
        if (other is null)
        {
            return false;
        }
        return Column.Equals(other.Column);
    }

    public override bool Equals(object? obj)
    {
        if (obj is ExcelPropertyInfo other)
        {
            return Equals(other);
        }
        return false;
    }

    public override int GetHashCode()
    {
        return Name.GetHashCode();
    }

    public override string ToString()
    {
        return $"{Name} ({Column})";
    }

    internal bool SetColumnId(ColumnId columnId)
    {
        if (_columnId.HasValue)
        {
            return false;
        }
        _columnId = columnId;
        return true;
    }

    private ColumnId GetColumnId()
    {
        if (_columnId.HasValue)
        {
            return _columnId.Value;
        }
        if (ProvidedColumnName != ColumnName.None)
        {
            _columnId = new ColumnId(ProvidedColumnName);
        }
        else
        {
            _columnId = new ColumnId(Index);
        }
        return _columnId.Value;
    }
}