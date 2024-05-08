namespace Eurostep.Excel
{
    public interface IExcelSheetRow
    {
        int RowNo { get; }

        string?[] GetValues();
    }
}