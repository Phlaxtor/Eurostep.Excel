namespace Eurostep.Excel;

public abstract class ExcelRowDefinition
{
    public string DetailsForLogging { get; internal set; }

    public uint RowId { get; internal set; }

    public string SheetName { get; init; }

    internal Dictionary<string, string> HeadingsWithColumnNames { get; set; }

    protected internal virtual uint? DescriptionRow => null;

    protected internal virtual uint FirstDataRow => 2;

    protected internal virtual uint HeaderRow => 1;
}