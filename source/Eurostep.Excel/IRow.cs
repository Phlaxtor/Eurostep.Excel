namespace Eurostep.Excel
{
    public interface IRow : ICollectionElement<ICell>
    {
        uint RowIndex { get; }
        ICell this[string column] { get; }
        ICell this[uint index] { get; }

        ICell? GetCellByColumnId(string column);

        ICell? GetCellByColumnIndex(uint index);

        IRowReader GetReader();

        ITableRow GetTableRow(IHeaderLookup lookup);

        string?[] GetValues();
    }
}