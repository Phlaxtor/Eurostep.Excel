namespace Eurostep.Excel;

internal sealed class FillStyleCache : ExcelStyleCache<IFillStyle, FillStyleValue>
{
    public FillStyleCache(ExcelWriter writer) : base(writer)
    {
    }

    protected override FillStyleValue Create(IFillStyle input)
    {
        ArgumentNullException.ThrowIfNull(input, nameof(input));
        return Writer.CreateFill(input);
    }
}