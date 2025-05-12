using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace Eurostep.Excel
{
    internal sealed class ExcelDocument : IExcelDocument
    {
        private readonly ExcelContext _context;
        private bool _disposed;

        public ExcelDocument(string path)
            : this(SpreadsheetDocument.Open(path, false))
        {
        }

        public ExcelDocument(Stream stream)
            : this(SpreadsheetDocument.Open(stream, false))
        {
        }

        internal ExcelDocument(string path, bool leaveOpen)
            : this(SpreadsheetDocument.Open(path, false), leaveOpen)
        {
        }

        internal ExcelDocument(Stream stream, bool leaveOpen)
            : this(SpreadsheetDocument.Open(stream, false), leaveOpen)
        {
        }

        private ExcelDocument(SpreadsheetDocument spreasheet, bool leaveOpen = true)
        {
            _context = new ExcelContext(spreasheet, leaveOpen);
        }

        public ISheet this[string name] => GetSheetByName(name) ?? throw new ApplicationException($"Can not find sheet with name '{name}'");
        public ISheet this[int index] => GetSheetByPosition(index) ?? throw new ApplicationException($"Can not find sheet at index '{index}'");

        public void Dispose()
        {
            if (_disposed) return;
            _context.Dispose();
            _disposed = true;
        }

        public IEnumerator<ISheet> GetEnumerator()
        {
            return new SheetEnumerator(_context.Sheets, _context);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return new SheetEnumerator(_context.Sheets, _context);
        }

        public ISheet? GetSheetById(string id)
        {
            int index = -1;
            foreach (Sheet sheet in _context.Sheets)
            {
                index++;
                if (sheet.Id is null) continue;
                if (sheet.Id.Value != id) continue;
                return new ExcelSheet(sheet, _context, index);
            }
            return default;
        }

        public ISheet? GetSheetByName(string name)
        {
            int index = -1;
            foreach (Sheet sheet in _context.Sheets)
            {
                index++;
                if (sheet.Name != name) continue;
                return new ExcelSheet(sheet, _context, index);
            }
            return default;
        }

        public ISheet? GetSheetByPosition(int position)
        {
            Sheet? sheet = (Sheet?)_context.Sheets.ElementAtOrDefault(position);
            if (sheet != null) return new ExcelSheet(sheet, _context, position);
            return default;
        }

        public ISheet? GetSheetBySheetId(uint sheetId)
        {
            int index = -1;
            foreach (Sheet sheet in _context.Sheets)
            {
                index++;
                if (sheet.SheetId is null) continue;
                if (sheet.SheetId.Value != sheetId) continue;
                return new ExcelSheet(sheet, _context, index);
            }
            return default;
        }
    }
}