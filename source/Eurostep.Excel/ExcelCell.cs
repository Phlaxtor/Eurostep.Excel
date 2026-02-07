using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    internal sealed class ExcelCell : ExcelElement, ICell
    {
        private readonly Cell _cell;
        private CellValues? _cellValues;
        private string? _column;
        private DataType? _dataType;
        private string? _reference;
        private uint? _row;
        private string? _value;

        public ExcelCell(Cell cell, ExcelContext context) : base(context)
        {
            _cell = cell;
        }

        public string CellReference => GetReference();
        public string Column => GetColumn();
        public uint Row => GetRow();
        public DataType Type => GetDataType();
        public string? Value => ReturnValue();

        public bool? GetBoolean()
        {
            if (TryGet(out bool value)) return value;
            return default;
        }

        public DateTime? GetDateTime()
        {
            if (TryGet(out DateTime value)) return value;
            return default;
        }

        public double? GetNumber()
        {
            if (TryGet(out double value)) return value;
            return default;
        }

        public string? GetText()
        {
            string? value = ReturnValue();
            if (string.IsNullOrEmpty(value)) return null;
            CellValues dataType = GetCellValuesType();

            if (dataType == CellValues.Boolean) return value.GetBooleanText();
            if (dataType == CellValues.Number) return value.GetNumberText();
            if (dataType == CellValues.Date) return value.GetDateTimeText();
            return value;
        }

        public override sealed string ToString()
        {
            return Value ?? string.Empty;
        }

        public bool TryGet(out bool value)
        {
            return bool.TryParse(Value, out value);
        }

        public bool TryGet(out double value)
        {
            return double.TryParse(Value, out value);
        }

        public bool TryGet(out DateTime value)
        {
            return DateTime.TryParse(Value, out value);
        }

        protected override int GetIndex()
        {
            return GetColumnIndex(Column);
        }

        protected override bool GetIsEmpty()
        {
            return string.IsNullOrWhiteSpace(Value);
        }

        private string? GetCellValue()
        {
            string? value = _cell.CellValue?.Text;
            if (string.IsNullOrEmpty(value)) return null;
            CellValues dataType = GetCellValuesType();

            if (dataType == CellValues.Error) return default;
            if (dataType == CellValues.SharedString) return Context.GetSharedString(value);
            if (dataType == CellValues.InlineString) return GetInlineString();
            return value;
        }

        private CellValues GetCellValuesType()
        {
            if (_cellValues.HasValue) return _cellValues.Value;
            _cellValues = _cell.DataType?.Value;
            if (_cellValues.HasValue) return _cellValues.Value;
            string? styleIndex = _cell.StyleIndex?.InnerText;
            _cellValues = Context.GetCellValues(styleIndex, CellValues.String);
            return _cellValues.Value;
        }

        private string GetColumn()
        {
            if (_column != null) return _column;
            (string Column, uint Row) reference = ParseReference(CellReference);
            _column = reference.Column;
            _row = reference.Row;
            return _column;
        }

        private DataType GetDataType()
        {
            if (_dataType.HasValue) return _dataType.Value;
            CellValues type = GetCellValuesType();
            if (type == CellValues.Boolean) _dataType = DataType.Boolean;
            else if (type == CellValues.Date) _dataType = DataType.DateTime;
            else if (type == CellValues.Error) _dataType = DataType.String;
            else if (type == CellValues.InlineString) _dataType = DataType.String;
            else if (type == CellValues.Number) _dataType = DataType.Number;
            else if (type == CellValues.SharedString) _dataType = DataType.String;
            else if (type == CellValues.String) _dataType = DataType.String;
            else _dataType = DataType.String;
            return _dataType.Value;
        }

        private string? GetInlineString()
        {
            string? text = _cell.InlineString?.Text?.Text;
            if (string.IsNullOrEmpty(text)) return null;
            return text;
        }

        private string GetReference()
        {
            if (_reference != null) return _reference;
            _reference = _cell.CellReference?.Value ?? throw new ArgumentNullException(nameof(CellReference));
            return _reference;
        }

        private uint GetRow()
        {
            if (_row.HasValue) return _row.Value;
            (string Column, uint Row) reference = ParseReference(CellReference);
            _column = reference.Column;
            _row = reference.Row;
            return _row.Value;
        }

        private string? ReturnValue()
        {
            if (_value != null) return _value;
            _value = GetCellValue();
            return _value;
        }
    }
}