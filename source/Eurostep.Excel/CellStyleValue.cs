using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    public readonly struct CellStyleValue
    {
        internal CellStyleValue(string name, uint value, uint? formatId, Alignment? alignment, BorderStyleValue? border, FillStyleValue? fill, FontStyleValue? font, NumberingFormatStyleValue? numberingFormat, bool? pivotButton, Protection? protection, bool? quotePrefix)
        {
            Alignment = alignment;
            Border = border;
            Fill = fill;
            Font = font;
            FormatId = formatId;
            Name = name;
            NumberingFormat = numberingFormat;
            PivotButton = pivotButton;
            Protection = protection;
            QuotePrefix = quotePrefix;
            Value = value;
        }

        public Alignment? Alignment { get; }
        public BorderStyleValue? Border { get; }
        public FillStyleValue? Fill { get; }
        public FontStyleValue? Font { get; }
        public uint? FormatId { get; }
        public bool HasAlignment => Alignment != null;
        public bool HasBorder => Border.HasValue;
        public bool HasFill => Fill.HasValue;
        public bool HasFont => Font.HasValue;
        public bool HasFormatId => FormatId.HasValue;
        public bool HasNumberingFormat => NumberingFormat.HasValue;
        public bool HasPivotButton => PivotButton.HasValue;
        public bool HasProtection => Protection != null;
        public bool HasQuotePrefix => QuotePrefix.HasValue;
        public string Name { get; }
        public NumberingFormatStyleValue? NumberingFormat { get; }
        public bool? PivotButton { get; }
        public Protection? Protection { get; }
        public bool? QuotePrefix { get; }
        public uint Value { get; }

        public static implicit operator uint(CellStyleValue value)
        {
            return value.Value;
        }
    }
}