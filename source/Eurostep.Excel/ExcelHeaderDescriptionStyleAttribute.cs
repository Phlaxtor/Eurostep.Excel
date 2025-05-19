namespace Eurostep.Excel;

public sealed class ExcelHeaderDescriptionStyleAttribute<T> : ExcelStylesheetAttribute<T>
    where T : ExcelStylesheetDefinition
{
    public ExcelHeaderDescriptionStyleAttribute() : base()
    {
    }

    public override ExcelStyleType StyleType => ExcelStyleType.HeaderDescription;
}