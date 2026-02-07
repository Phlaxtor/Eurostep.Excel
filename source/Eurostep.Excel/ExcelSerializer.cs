namespace Eurostep.Excel;

public sealed class ExcelSerializer
{
    private readonly Dictionary<Type, ExcelRowInfo> _typeToRowInfo = [];

    public ExcelSerializer(Type type)
    {
        _typeToRowInfo[type] = new ExcelRowInfo(type);
    }

    public ExcelSerializer(IEnumerable<Type> types)
    {
        foreach (Type type in types)
        {
            _typeToRowInfo[type] = new ExcelRowInfo(type);
        }
    }

    public IReadOnlyCollection<T> Deserialize<T>(ExcelReader reader, ExcelSerializationSettings settings)
    {
        List<T> collection = [];
        return collection;
    }

    public void Serialize<T>(ExcelWriter writer, IEnumerable<T> collection, ExcelSerializationSettings settings)
    {
        foreach (T item in collection)
        {
            Serialize(writer, item);
        }
    }

    private void Serialize<T>(ExcelWriter writer, T item)
    {
    }

    private void Initialize(ExcelWriter writer, ExcelSerializationSettings settings, Type type)
    {
        if (_typeToRowInfo.TryGetValue(type, out ExcelRowInfo? rowInfo) == false)
        {
            throw new ApplicationException();
        }
        writer.SetCurrentTab(settings.SheetName);
        if (settings.UseHeaders)
        {
            ExcelPropertyInfo[] properties = rowInfo.GetProperties();
            WriteHeader(writer, properties);
        }
    }

    private void WriteHeader(ExcelWriter writer, ExcelPropertyInfo[] properties)
    {
        HeaderBuilder builder = writer.AddHeaders();
        foreach (ExcelPropertyInfo item in properties)
        {
            builder.New(item.Name, item.Width);
        }
        builder.Build();
    }
}