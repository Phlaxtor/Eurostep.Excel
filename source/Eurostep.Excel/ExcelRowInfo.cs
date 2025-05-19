using System.Diagnostics.CodeAnalysis;

namespace Eurostep.Excel;

public sealed class ExcelRowInfo
{
    private readonly ExcelPropertyInfo[] _properties;
    private readonly Dictionary<string, ExcelPropertyInfo> _propertyLookup;
    private readonly Type _type;

    public ExcelRowInfo(Type type)
    {
        _type = type ?? throw new ArgumentNullException(nameof(type));
        _propertyLookup = new Dictionary<string, ExcelPropertyInfo>();
        var properties = type.GetExcelProperties();
        _properties = new ExcelPropertyInfo[properties.Count];
        foreach (var property in properties)
        {
            _properties[property.Index] = property;
            _propertyLookup[property.Id] = property;
        }

        foreach (ExcelAttribute attribute in type.GetAttributes<ExcelAttribute>())
        {
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
    }

    public ExcelStylesheetDefinition? CellStyle { get; }

    [MemberNotNullWhen(true, nameof(CellStyle))]
    public bool HasGeneralCellStyle { get; }

    [MemberNotNullWhen(true, nameof(HeaderDescriptionStyle))]
    public bool HasGeneralHeaderDescriptionStyle { get; }

    [MemberNotNullWhen(true, nameof(HeaderStyle))]
    public bool HasGeneralHeaderStyle { get; }

    public ExcelStylesheetDefinition? HeaderDescriptionStyle { get; }

    public ExcelStylesheetDefinition? HeaderStyle { get; }
}