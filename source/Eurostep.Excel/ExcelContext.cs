using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    internal sealed class ExcelContext : IDisposable
    {
        private readonly CellFormat[] _cellFormats;
        private readonly SharedStringItem[] _sharedStrings;
        private bool _cellFormatsInitialized = false;
        private bool _disposed = false;
        private bool _sharedStringsInitialized = false;

        public ExcelContext(SpreadsheetDocument spreasheet, bool leaveOpen)
        {
            Spreasheet = spreasheet;
            LeaveOpen = leaveOpen;
            _cellFormats = new CellFormat[CellFormatCount];
            _sharedStrings = new SharedStringItem[SharedStringCount];
        }

        public uint CellFormatCount => CellFormats?.Count ?? 0;
        public CellFormats? CellFormats => Stylesheet?.CellFormats;
        public bool LeaveOpen { get; }
        public uint SharedStringCount => SharedStringTable?.Count ?? 0;
        public SharedStringTable? SharedStringTable => SharedStringTablePart?.SharedStringTable;
        public SharedStringTablePart? SharedStringTablePart => WorkbookPart.SharedStringTablePart;
        public Sheets Sheets => Workbook.Sheets ?? throw new ArgumentNullException(nameof(Sheets));
        public SpreadsheetDocument Spreasheet { get; }
        public Stylesheet? Stylesheet => WorkbookStylesPart?.Stylesheet;
        public Workbook Workbook => WorkbookPart.Workbook ?? throw new ArgumentNullException(nameof(Workbook));
        public WorkbookPart WorkbookPart => Spreasheet.WorkbookPart ?? throw new ArgumentNullException(nameof(WorkbookPart));
        public WorkbookStylesPart? WorkbookStylesPart => WorkbookPart.WorkbookStylesPart;

        public void Dispose()
        {
            if (_disposed) return;
            Spreasheet.Dispose();
            _disposed = true;
        }

        public CellFormat? GetCellFormat(int index)
        {
            if (_cellFormats.Length <= index) return null;
            CellFormatsLookupInit();
            return _cellFormats[index];
        }

        public CellFormat? GetCellFormat(string? index)
        {
            if (int.TryParse(index, out int value) == false) return null;
            return GetCellFormat(value);
        }

        public CellValues GetCellValues(string? index, CellValues defaultValue = CellValues.String)
        {
            CellFormat? cellFormat = GetCellFormat(index);
            if (cellFormat == null) return defaultValue;
            uint? numberFormat = cellFormat.NumberFormatId?.Value;
            if (numberFormat.HasValue == false) return defaultValue;
            switch (numberFormat.Value)
            {
                case 0: return CellValues.String;
                case 1: return CellValues.String;
                case 2: return CellValues.Number;
                case 3: return CellValues.Number;
                case 4: return CellValues.Number;
                case 9: return CellValues.String;
                case 10: return CellValues.Number;
                case 11: return CellValues.Number;
                case 12: return CellValues.String;
                case 13: return CellValues.String;
                case 14: return CellValues.Date;
                case 15: return CellValues.Date;
                case 16: return CellValues.Date;
                case 17: return CellValues.Date;
                case 18: return CellValues.Date;
                case 19: return CellValues.Date;
                case 20: return CellValues.Date;
                case 21: return CellValues.Date;
                case 22: return CellValues.Date;
                case 37: return CellValues.Number;
                case 38: return CellValues.Number;
                case 39: return CellValues.Number;
                case 40: return CellValues.Number;
                case 45: return CellValues.Date;
                case 46: return CellValues.Date;
                case 47: return CellValues.Date;
                case 48: return CellValues.Number;
                case 49: return CellValues.String;
                default: return defaultValue;
            }
        }

        public T GetPart<T>(string id) where T : OpenXmlPart
        {
            OpenXmlPart part = WorkbookPart.GetPartById(id);
            if (part is T p) return p;
            throw new ApplicationException($"There is no {typeof(T)} with id '{id}'");
        }

        public string? GetSharedString(int reference)
        {
            if (_sharedStrings.Length <= reference) return null;
            SharedStringsLookupInit();
            SharedStringItem value = _sharedStrings[reference];
            return value?.InnerText;
        }

        public string? GetSharedString(string? reference)
        {
            if (string.IsNullOrEmpty(reference)) return null;
            if (int.TryParse(reference, out int r)) return GetSharedString(r);
            throw new ArgumentException($"Reference '{reference}' is not a valid reference", nameof(reference));
        }

        public WorksheetPart GetWorksheetPart(string id)
        {
            return GetPart<WorksheetPart>(id);
        }

        private void CellFormatsLookupInit()
        {
            if (_cellFormatsInitialized) return;
            if (CellFormats == null) return;
            int index = 0;
            foreach (DocumentFormat.OpenXml.OpenXmlElement f in CellFormats)
            {
                if (f is not CellFormat format) continue;
                _cellFormats[index] = format;
                index++;
            }
            _cellFormatsInitialized = true;
        }

        private void SharedStringsLookupInit()
        {
            if (_sharedStringsInitialized) return;
            if (SharedStringTable == null) return;
            int index = 0;
            foreach (DocumentFormat.OpenXml.OpenXmlElement i in SharedStringTable)
            {
                if (i is not SharedStringItem item) continue;
                _sharedStrings[index] = item;
                index++;
            }
            _sharedStringsInitialized = true;
        }
    }
}