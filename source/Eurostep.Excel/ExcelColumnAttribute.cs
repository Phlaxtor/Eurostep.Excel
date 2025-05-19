namespace Eurostep.Excel;

[AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
public sealed class ExcelColumnAttribute : ExcelAttribute
{
    public const double DefaultWidth = 10;

    public ExcelColumnAttribute(ColumnName column, double width = DefaultWidth)
    {
        Column = column;
        Width = width < DefaultWidth ? DefaultWidth : width;
    }

    public ColumnName Column { get; }

    public double Width { get; }
}