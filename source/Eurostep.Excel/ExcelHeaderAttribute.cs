namespace Eurostep.Excel;

[AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
public sealed class ExcelHeaderAttribute : ExcelAttribute
{
    public ExcelHeaderAttribute(string heading, string? description = null)
    {
        Description = description;
        Name = heading;
    }

    public string? Description { get; }

    public string Name { get; }
}