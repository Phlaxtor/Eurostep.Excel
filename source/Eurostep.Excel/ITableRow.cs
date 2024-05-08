namespace Eurostep.Excel
{
    public interface ITableRow : ICollectionElement<ICell>
    {
        string? this[string header] { get; }

        bool? GetBoolean(string header);

        ICell? GetCell(string header);

        DateTime? GetDateTime(string header);

        double? GetNumber(string header);

        string? GetText(string header);
    }
}