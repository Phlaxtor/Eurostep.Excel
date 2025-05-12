namespace Eurostep.Excel;

internal sealed class NumberingFormatStyleCache
{
    private readonly Dictionary<string, NumberingFormatStyleValue> _cache;
    private readonly Type _type = typeof(INumberingFormatStyle);
    private readonly ExcelWriter _writer;
    private uint _numberFormatId = 0;

    public NumberingFormatStyleCache(ExcelWriter writer)
    {
        _cache = [];
        _writer = writer;
    }

    public NumberingFormatStyleValue Get(INumberingFormatStyle input)
    {
        ArgumentNullException.ThrowIfNull(input, nameof(input));
        ArgumentException.ThrowIfNullOrWhiteSpace(input.FormatCode, nameof(input));
        return Get(input.FormatCode);
    }

    public NumberingFormatStyleValue Get<T>() where T : INumberingFormatStyle
    {
        return Get(typeof(T));
    }

    public NumberingFormatStyleValue Get(Type type)
    {
        ArgumentNullException.ThrowIfNull(type, nameof(type));
        if (_type.IsAssignableFrom(type) == false)
        {
            throw new ArgumentException();
        }

        INumberingFormatStyle? instance = (INumberingFormatStyle?)Activator.CreateInstance(type);
        ArgumentNullException.ThrowIfNull(instance, nameof(instance));
        ArgumentException.ThrowIfNullOrWhiteSpace(instance.FormatCode, nameof(instance));
        return Get(instance.FormatCode);
    }

    public NumberingFormatStyleValue Get(string formatCode)
    {
        ArgumentNullException.ThrowIfNullOrWhiteSpace(formatCode, nameof(formatCode));
        if (_cache.TryGetValue(formatCode, out NumberingFormatStyleValue value))
        {
            return value;
        }

        _numberFormatId++;
        value = _writer.CreateNumberingFormat(_numberFormatId, formatCode);
        _cache[formatCode] = value;
        return value;
    }

    public NumberingFormatStyleValue? GetOrDefault(INumberingFormatStyle? input)
    {
        if (input is null)
        {
            return default;
        }

        return Get(input);
    }
}