namespace Eurostep.Excel;

public sealed class ExcelHeaderStyleAttribute<T> : ExcelStylesheetAttribute<T>
    where T : ExcelStylesheetDefinition
{
    public ExcelHeaderStyleAttribute() : base()
    {
    }

    public override ExcelStyleType StyleType => ExcelStyleType.Header;
}