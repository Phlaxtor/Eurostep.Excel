namespace Eurostep.Excel;

internal sealed class FontStyleCache : ExcelStyleCache<IFontStyle, FontStyleValue>
{
    public FontStyleCache(ExcelWriter writer) : base(writer)
    {
    }

    protected override FontStyleValue Create(IFontStyle input)
    {
        ArgumentNullException.ThrowIfNull(input, nameof(input));
        return Writer.CreateFont(input);
    }
}