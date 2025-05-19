using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace Eurostep.Excel;

public sealed class ExcelPropertyInfo
{
    private readonly PropertyInfo _value;

    public ExcelPropertyInfo(PropertyInfo value, int index)
    {
        Index = index;
        _value = value ?? throw new ArgumentNullException(nameof(value));
        foreach (ExcelAttribute attribute in value.GetAttributes<ExcelAttribute>())
        {
            if (attribute is ExcelHeaderAttribute header)
            {
                HeaderName = header.Name;
                HeaderDescription = header.Description;
                continue;
            }
            if (attribute is ExcelColumnAttribute column)
            {
                Column = column.Column;
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

        Id = (HeaderName ?? PropertyName).GetToUpperWithoutWhiteSpace();
    }

    public ExcelStylesheetDefinition? CellStyle { get; }

    public ColumnName Column { get; }

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

    public int Index { get; }

    public string PropertyName => _value.Name;

    public bool Required { get; }

    public double Width { get; }
}