using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    public struct CellStyle
    {
        private readonly Alignment? _alignment;
        private readonly uint? _border;
        private readonly uint _cellFormat;
        private readonly uint? _fill;
        private readonly uint? _font;
        private readonly uint? _formatId;
        private readonly string _name;
        private readonly uint? _numberFormatId;
        private readonly uint? _numberingFormat;
        private readonly bool? _pivotButton;
        private readonly Protection? _protection;
        private readonly bool? _quotePrefix;

        internal CellStyle(string name, uint cellFormat, uint? formatId, Alignment? alignment, uint? border, uint? fill, uint? font, uint? numberFormatId, uint? numberingFormat, bool? pivotButton, Protection? protection, bool? quotePrefix)
        {
            _alignment = alignment;
            _border = border;
            _cellFormat = cellFormat;
            _fill = fill;
            _font = font;
            _formatId = formatId;
            _name = name;
            _numberFormatId = numberFormatId;
            _numberingFormat = numberingFormat;
            _pivotButton = pivotButton;
            _protection = protection;
            _quotePrefix = quotePrefix;
        }

        public Alignment? Alignment => _alignment;
        public uint? Border => _border;
        public uint? Fill => _fill;
        public uint? Font => _font;
        public uint? FormatId => _formatId;
        public bool HasAlignment => _alignment != null;
        public bool HasBorder => _border.HasValue;
        public bool HasFill => _fill.HasValue;
        public bool HasFont => _font.HasValue;
        public bool HasFormatId => _formatId.HasValue;
        public bool HasNumberingFormat => _numberFormatId.HasValue;
        public bool HasPivotButton => _pivotButton.HasValue;
        public bool HasProtection => _protection != null;
        public bool HasQuotePrefix => _quotePrefix.HasValue;
        public string Name => _name;
        public uint? NumberFormatId => _numberFormatId;
        public uint? NumberingFormat => _numberingFormat;
        public bool? PivotButton => _pivotButton;
        public Protection? Protection => _protection;
        public bool? QuotePrefix => _quotePrefix;
        public uint Value => _cellFormat;

        public static implicit operator uint(CellStyle c) => c.Value;
    }
}