namespace Eurostep.Excel;

internal sealed class BorderStyleCache : ExcelStyleCache<IBorderStyle, BorderStyleValue>
{
    public BorderStyleCache(ExcelWriter writer) : base(writer)
    {
    }

    protected override BorderStyleValue Create(IBorderStyle input)
    {
        ArgumentNullException.ThrowIfNull(input, nameof(input));
        return Writer.CreateBorder(input);
    }
}