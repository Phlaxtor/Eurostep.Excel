namespace Eurostep.Excel;

public class ExcelColumnAttribute : ExcelAttribute
{
    public ExcelColumnAttribute(string heading)
    {
        Heading = heading;
    }

    public ExcelColumnAttribute(string heading, string column, string description)
    {
        Heading = heading;
        Column = column;
        Description = description;
    }

    internal virtual string Column { get; set; }

    internal virtual string Heading { get; private set; }

    internal virtual string Description { get; private set; }
}