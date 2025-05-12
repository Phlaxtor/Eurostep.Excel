namespace Eurostep.Excel;

public sealed class ExcelCellStyleAttribute<T> : ExcelStylesheetAttribute<T>
    where T : ExcelStylesheetDefinition
{
    public ExcelCellStyleAttribute() : base()
    {
    }
}