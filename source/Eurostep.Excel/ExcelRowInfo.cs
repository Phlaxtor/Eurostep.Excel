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
            _propertyLookup.Add(property.Id, property);
        }

        foreach (ExcelAttribute attribute in type.GetAttributes<ExcelAttribute>())
        {
            if (attribute is ExcelStylesheetAttribute style)
            {
                switch (style.StyleType)
                {
                    case ExcelStyleType.Cell:
                        GeneralCellStyle = style.GetStylesheetDefinition();
                        break;

                    case ExcelStyleType.Header:
                        GeneralHeaderStyle = style.GetStylesheetDefinition();
                        break;

                    case ExcelStyleType.HeaderDescription:
                        GeneralHeaderDescriptionStyle = style.GetStylesheetDefinition();
                        break;
                }
                continue;
            }
        }
    }

    public ExcelStylesheetDefinition? GeneralCellStyle { get; }

    public ExcelStylesheetDefinition? GeneralHeaderDescriptionStyle { get; }

    public ExcelStylesheetDefinition? GeneralHeaderStyle { get; }

    [MemberNotNullWhen(true, nameof(GeneralCellStyle))]
    public bool HasGeneralCellStyle { get; }

    [MemberNotNullWhen(true, nameof(GeneralHeaderDescriptionStyle))]
    public bool HasGeneralHeaderDescriptionStyle { get; }

    [MemberNotNullWhen(true, nameof(GeneralHeaderStyle))]
    public bool HasGeneralHeaderStyle { get; }

    internal ExcelPropertyInfo[] GetProperties()
    {
        var properties = new List<ExcelPropertyInfo>(_propertyLookup.Values);
        properties.Sort();
        return properties.ToArray();
    }

    internal bool SetHeader(HeaderId header)
    {
        if (_propertyLookup.TryGetValue(header.Id, out ExcelPropertyInfo? info))
        {
            return info.SetColumnId(header.Column);
        }
        return false;
    }
}