namespace Eurostep.Excel
{
    public interface ISheetWriter : IDisposable
    {
        int RowCount { get; }
        string SheetName { get; }

        void AddHeaders(IPresentationColumn[] headers);

        HeaderBuilder AddHeaders();

        void AddIntegerValidation(ColumnId applyColumn, uint applyRowStart, int minValue, int maxValue, string? errorTitle = null);

        void AddMandatoryCellCheck(ColumnId mandatoryColumn, uint rowStart, params ColumnId[] checkColumns);

        void AddRangeValidation(ColumnId applyColumn, uint applyRowStart, ColumnId valuesColumn, uint valuesCount, string validationSheetId, string? errorTitle = null);

        void AddRow(params string[] values);

        void AddRow(params ICellValue[] values);

        RowBuilder AddRow();

        int AddRowCount();

        Task AddRows(IAsyncEnumerable<IExcelSheetRow> rows);

        void AddVerticalColumn(string header, CellStyle? styleIndex, params string[] values);

        void Close();

        void EndSheet(bool addTable);
    }
}