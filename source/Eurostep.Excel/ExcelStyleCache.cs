namespace Eurostep.Excel;

internal abstract class ExcelStyleCache<TInput, TOutput>
{
    private readonly Dictionary<Type, TOutput> _cache;
    private readonly Type _type = typeof(TInput);
    private readonly ExcelWriter _writer;

    protected ExcelStyleCache(ExcelWriter writer)
    {
        _cache = [];
        _writer = writer;
    }

    protected ExcelWriter Writer => _writer;

    public TOutput Get(TInput input)
    {
        ArgumentNullException.ThrowIfNull(input, nameof(input));
        return Get(input.GetType());
    }

    public TOutput Get<T>() where T : TInput
    {
        return Get(typeof(T));
    }

    public TOutput Get(Type type)
    {
        ArgumentNullException.ThrowIfNull(type, nameof(type));
        if (_type.IsAssignableFrom(type) == false)
        {
            throw new ArgumentException();
        }
        if (_cache.TryGetValue(type, out TOutput? value) == false)
        {
            TInput? instance = (TInput?)Activator.CreateInstance(type);
            ArgumentNullException.ThrowIfNull(instance, nameof(instance));
            _cache[type] = value = Create(instance);
        }
        return value;
    }

    public TOutput? GetOrDefault(TInput? input)
    {
        if (input is null)
        {
            return default;
        }

        return Get(input);
    }

    protected abstract TOutput Create(TInput input);
}