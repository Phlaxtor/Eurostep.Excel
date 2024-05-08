namespace Eurostep.Excel
{
    public interface IExcelDocument : IEnumerable<ISheet>, IDisposable
    {
        ISheet this[string name] { get; }
        ISheet this[int index] { get; }

        ISheet? GetSheetById(string id);

        ISheet? GetSheetByPosition(int position);

        ISheet? GetSheetByName(string name);

        ISheet? GetSheetBySheetId(uint sheetId);
    }
}