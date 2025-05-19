using System.Collections;

namespace Eurostep.Excel;

public abstract class ExcelSerializer
{
    private readonly ExcelRowInfo _rowInfo;
    private readonly Type _type;

    protected ExcelSerializer(Type type)
    {
        _type = type;
        _rowInfo = new ExcelRowInfo(type);
    }

    protected IEnumerable<T> DeserializeCollection<T>(ExcelReader reader)
    {
        ArgumentNullException.ThrowIfNull(reader, nameof(reader));
        return Enumerable.Empty<T>();
    }

    protected IEnumerable DeserializeCollection(ExcelReader reader)
    {
        ArgumentNullException.ThrowIfNull(reader, nameof(reader));
        return default;
    }

    protected void Initialize(ExcelReader reader)
    {
        ArgumentNullException.ThrowIfNull(reader, nameof(reader));
    }

    protected void Serialize(ExcelWriter writer, IEnumerable collection)
    {
        ArgumentNullException.ThrowIfNull(writer, nameof(writer));
        ArgumentNullException.ThrowIfNull(collection, nameof(collection));
        foreach (object? item in collection)
        {
            Serialize(writer, item);
        }
    }

    protected void Serialize(ExcelWriter writer, object item)
    {
        ArgumentNullException.ThrowIfNull(writer, nameof(writer));
        ArgumentNullException.ThrowIfNull(item, nameof(item));
    }
}

public sealed class ExcelSerializer<T> : ExcelSerializer
{
    public ExcelSerializer() : base(typeof(T))
    {
    }

    public IReadOnlyCollection<T> Deserialize(ExcelReader reader)
    {
        List<T> collection = [.. DeserializeCollection<T>(reader)];
        return collection;
    }

    public void Serialize(ExcelWriter writer, IEnumerable<T> collection)
    {
        foreach (T item in collection)
        {
            Serialize(writer, item);
        }
    }
}