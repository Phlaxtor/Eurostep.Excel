using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace Eurostep.Excel
{
    internal sealed class ExcelSheet : ExcelElement, ISheet
    {
        private readonly int _index;
        private readonly SheetData _sheetData;
        private int? _length;

        public ExcelSheet(Sheet sheet, ExcelContext context, int index) : base(context)
        {
            Id = sheet.Id?.Value ?? throw new ArgumentNullException(nameof(Id));
            Name = sheet.Name?.Value ?? throw new ArgumentNullException(nameof(Name));
            SheetId = sheet.SheetId?.Value ?? throw new ArgumentNullException(nameof(SheetId));

            WorksheetPart worksheetPart = context.GetWorksheetPart(Id);
            _sheetData = GetSheetData(worksheetPart.Worksheet);
            _index = index;
        }

        public string Id { get; }
        public int Length => GetLength();
        public string Name { get; }
        public uint SheetId { get; }
        public IRow this[uint index] => GetRowByRowIndex(index);

        public IEnumerator<IRow> GetEnumerator()
        {
            return new RowEnumerator(_sheetData, Context);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return new RowEnumerator(_sheetData, Context);
        }

        public ISheetReader GetReader()
        {
            return new SheetReader(this, Context);
        }

        public IRow GetRowByRowIndex(uint rowIndex)
        {
            foreach (DocumentFormat.OpenXml.OpenXmlElement element in _sheetData)
            {
                if (element is not Row row) continue;
                if (row.RowIndex is null) continue;
                if (row.RowIndex.Value != rowIndex) continue;
                return new ExcelRow(row, Context);
            }
            throw new ApplicationException($"Can not find row with row index '{rowIndex}'");
        }

        protected override int GetIndex()
        {
            return _index;
        }

        protected override bool GetIsEmpty()
        {
            using IEnumerator<IRow> enumerator = GetEnumerator();
            while (enumerator.MoveNext())
            {
                if (enumerator.Current.IsEmpty == false) return false;
            }
            return true;
        }

        private int GetLength()
        {
            if (_length.HasValue) return _length.Value;
            if (_sheetData.HasChildren == false) return 0;
            if (_sheetData.LastChild is null) return 0;
            if (_sheetData.LastChild is not Row l) throw new ApplicationException("LastChild on SheetData is not a Row");
            uint i = l.RowIndex?.Value ?? throw new ArgumentNullException(nameof(Row.RowIndex));
            _length = (int)i;
            return _length.Value;
        }

        private IRow GetRowByPosition(int position)
        {
            int i = -1;
            foreach (DocumentFormat.OpenXml.OpenXmlElement element in _sheetData)
            {
                if (element is not Row row) continue;
                i++;
                if (i < position) continue;
                return new ExcelRow(row, Context);
            }
            throw new ApplicationException($"Can not find row at position '{position}'");
        }

        private SheetData GetSheetData(Worksheet worksheet)
        {
            foreach (DocumentFormat.OpenXml.OpenXmlElement element in worksheet.ChildElements)
            {
                if (element is SheetData e) return e;
            }
            throw new ArgumentException($"{nameof(Worksheet)} does not have any valid {nameof(SheetData)}", nameof(worksheet));
        }
    }
}