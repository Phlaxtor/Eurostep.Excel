namespace Eurostep.Excel
{
    public interface ISheet : ICollectionElement<IRow>
    {
        string Id { get; }
        string Name { get; }
        uint SheetId { get; }
        IRow this[uint index] { get; }

        ISheetReader GetReader();

        IRow? GetRowByRowIndex(uint rowIndex);
    }
}