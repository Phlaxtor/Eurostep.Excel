namespace Eurostep.Excel;

public abstract class ExcelStylesheetAttribute<T> : ExcelStylesheetAttribute
    where T : ExcelStylesheetDefinition
{
    protected ExcelStylesheetAttribute() : base()
    {
    }

    protected override Type GetDefinitionType()
    {
        return typeof(T);
    }
}

[AttributeUsage(AttributeTargets.Class | AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
public abstract class ExcelStylesheetAttribute : ExcelAttribute
{
    protected ExcelStylesheetAttribute() : base()
    {
    }

    public abstract ExcelStyleType StyleType { get; }

    internal ExcelStylesheetDefinition GetStylesheetDefinition()
    {
        object? definition = Activator.CreateInstance(GetDefinitionType());
        ArgumentNullException.ThrowIfNull(definition, nameof(definition));
        if (definition is ExcelStylesheetDefinition value)
        {
            return value;
        }
        throw new ApplicationException($"Not of type {nameof(ExcelStylesheetDefinition)}");
    }

    internal Type GetStylesheetDefinitionType()
    {
        return GetDefinitionType();
    }

    protected abstract Type GetDefinitionType();
}