using System.Drawing;
using System.Text;

namespace Eurostep.Excel
{
    public sealed class SheetWriter : ISheetWriter, IDisposable
    {
        public const string ValidationSheetId = "Values";
        private readonly string _sheetName;
        private readonly ExcelWriter _writer;

        internal SheetWriter(ExcelWriter writer, string sheetName)
        {
            _writer = writer;
            _sheetName = sheetName;
        }

        public int RowCount { get; private set; }

        public string SheetName => _sheetName;

        public static ISheetWriter GetClient(Stream stream, string sheetName)
        {
            ExcelWriter writer = new ExcelWriter(stream);
            writer.AddNewTab(sheetName);
            return new SheetWriter(writer, sheetName);
        }

        public void AddHeaders(IPresentationColumn[] headers)
        {
            RowCount++;
            _writer.SetCurrentTab(SheetName);
            _writer.AddHeaders(headers);
        }

        public HeaderBuilder AddHeaders()
        {
            return new HeaderBuilder(_writer, SheetName);
        }

        public void AddIntegerValidation(ColumnId applyColumn, uint applyRowStart, int minValue, int maxValue, string? errorTitle = null)
        {
            _writer.SetCurrentTab(SheetName);
            ColumnRange validationRange = new ColumnRange(applyColumn, applyRowStart, ColumnRange.RowMax);
            string errorText = $"Value must be a number between {minValue} - {maxValue}";
            _writer.AddDataValidationWhole(validationRange, minValue, maxValue, errorText, errorTitle);
        }

        public void AddMandatoryCellCheck(ColumnId mandatoryColumn, uint rowStart, params ColumnId[] checkColumns)
        {
            _writer.SetCurrentTab(SheetName);
            ColumnRange range = new ColumnRange(mandatoryColumn, rowStart);
            StringBuilder condition = new StringBuilder($"AND({mandatoryColumn}{rowStart}=\"\", OR(");
            string delimiter = string.Empty;
            foreach (ColumnId c in checkColumns)
            {
                condition.Append(delimiter);
                condition.Append($"{c}{rowStart}<>\"\"");
                delimiter = ", ";
            }
            condition.Append("))");
            _writer.AddMandatoryCellCheck(range, condition.ToString(), KnownColor.Red);
        }

        public void AddRangeValidation(ColumnId applyColumn, uint applyRowStart, ColumnId valuesColumn, uint valuesCount, string validationSheetId, string? errorTitle = null)
        {
            _writer.SetCurrentTab(SheetName);
            ColumnRange allowedValues = new ColumnRange(valuesColumn, 2, 2 + valuesCount, validationSheetId);
            ColumnRange validationRange = new ColumnRange(applyColumn, applyRowStart, ColumnRange.RowMax);
            string errorText = $"Value not allowed";
            _writer.AddDataValidationList(validationRange, allowedValues, errorText, errorTitle);
        }

        public void AddRow(params string[] values)
        {
            RowCount++;
            _writer.SetCurrentTab(SheetName);
            _writer.AddNewRow(values);
        }

        public void AddRow(params ICellValue[] values)
        {
            RowCount++;
            _writer.SetCurrentTab(SheetName);
            _writer.AddNewRow(values);
        }

        public RowBuilder AddRow()
        {
            return new RowBuilder(this);
        }

        public int AddRowCount()
        {
            RowCount++;
            return RowCount;
        }

        public async Task AddRows(IAsyncEnumerable<IExcelSheetRow> rows)
        {
            _writer.SetCurrentTab(SheetName);
            await foreach (IExcelSheetRow row in rows)
            {
                string?[] values = row.GetValues();
                _writer.AddNewRow(values);
            }
        }

        public void AddVerticalColumn(string header, CellStyleValue? style, params string[] values)
        {
            _writer.SetCurrentTab(SheetName);
            RowCount++;
            _writer.AddVerticalColumn(header, style, values);
        }

        public void Close()
        {
            _writer.Close();
        }

        public void Dispose()
        {
            _writer.Dispose();
        }

        public void EndSheet(bool addTable)
        {
            _writer.CloseTab(SheetName, addTable);
        }
    }
}