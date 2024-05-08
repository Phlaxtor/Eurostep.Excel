namespace Eurostep.Excel
{
    public interface ICell : IElement
    {
        string Column { get; }
        uint Row { get; }
        DataType Type { get; }
        string? Value { get; }

        bool? GetBoolean();

        DateTime? GetDateTime();

        double? GetNumber();

        string? GetText();

        bool TryGet(out bool value);

        bool TryGet(out double value);

        bool TryGet(out DateTime value);
    }
}