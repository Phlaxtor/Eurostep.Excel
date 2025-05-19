using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    public sealed class CellStyleBuilder
    {
        private readonly ExcelWriter _excel;
        private readonly uint? _formatId;
        private Alignment? _alignment;
        private BorderStyleValue? _border;
        private CellStyleValue? _cellFormat;
        private FillStyleValue? _fill;
        private FontStyleValue? _font;
        private NumberingFormatStyleValue? _numberingFormat;
        private bool? _pivotButton;
        private Protection? _protection;
        private bool? _quotePrefix;

        internal CellStyleBuilder(ExcelWriter excel)
        {
            _excel = excel;
        }

        public bool HasAlignment => _alignment != null;
        public bool HasBorder => _border.HasValue;
        public bool HasCellFormat => _cellFormat.HasValue;
        public bool HasFill => _fill.HasValue;
        public bool HasFont => _font.HasValue;
        public bool HasFormatId => _formatId.HasValue;
        public bool HasNumberingFormat => _numberingFormat.HasValue;
        public bool HasPivotButton => _pivotButton.HasValue;
        public bool HasProtection => _protection != null;
        public bool HasQuotePrefix => _quotePrefix.HasValue;

        public CellStyleValue Build()
        {
            if (_cellFormat.HasValue) throw new ApplicationException($"Cell format style has alread been set");
            _cellFormat = _excel.NewCellStyle(_numberingFormat, _formatId, _alignment, _font, _border, _fill, _protection, _pivotButton, _quotePrefix);
            return _cellFormat.Value;
        }

        public Alignment? GetAlignment()
        {
            return _alignment;
        }

        public BorderStyleValue? GetBorder()
        {
            return _border;
        }

        public CellStyleValue? GetCellFormat()
        {
            return _cellFormat;
        }

        public FillStyleValue? GetFill()
        {
            return _fill;
        }

        public FontStyleValue? GetFont()
        {
            return _font;
        }

        public uint? GetFormatId()
        {
            return _formatId;
        }

        public NumberingFormatStyleValue? GetNumberingFormat()
        {
            return _numberingFormat;
        }

        public bool? GetPivotButton()
        {
            return _pivotButton;
        }

        public Protection? GetProtection()
        {
            return _protection;
        }

        public bool? GetQuotePrefix()
        {
            return _quotePrefix;
        }

        public CellStyleBuilder SetAlignment(HorizontalAlignmentValues horizontal, VerticalAlignmentValues vertical, uint indent = 0, int relativeIndent = 0, bool shrinkToFit = false, bool wrapText = false, uint textRotation = 0, string? mergeCell = null, uint readingOrder = 0, bool justifyLastLine = false)
        {
            if (_alignment != null) throw new ApplicationException($"Alignment style has alread been set");
            _alignment = _excel.GetAlignment(horizontal, vertical, indent, relativeIndent, shrinkToFit, wrapText, textRotation, mergeCell, readingOrder, justifyLastLine);
            return this;
        }

        public CellStyleBuilder SetBorder(BorderPart? top = null, BorderPart? right = null, BorderPart? bottom = null, BorderPart? left = null)
        {
            if (_border.HasValue) throw new ApplicationException($"Border style has alread been set");
            _border = _excel.CreateBorder(top, right, bottom, left);
            return this;
        }

        public CellStyleBuilder SetFill(PatternValues? patternType = null, GeneralColor? fgColor = null, GeneralColor? bgColor = null, GradientValues? gradientType = null, double degree = 0, double top = 0, double bottom = 0, double right = 0, double left = 0)
        {
            if (_fill.HasValue) throw new ApplicationException($"Fill style has alread been set");
            _fill = _excel.CreateFill(patternType, fgColor, bgColor, gradientType, degree, top, bottom, right, left);
            return this;
        }

        public CellStyleBuilder SetFont(string? name = "Calibri", double? sz = null, bool? b = null, bool? i = null, UnderlineValues? u = null, GeneralColor? color = null, VerticalAlignmentRunValues? vertAlig = null, bool? strike = null, bool? condense = null, bool? extend = null, bool? shadow = null)
        {
            if (_font.HasValue) throw new ApplicationException($"Font style has alread been set");
            _font = _excel.CreateFont(name, sz, b, i, u, color, vertAlig, strike, condense, extend, shadow);
            return this;
        }

        public CellStyleBuilder SetNumberingFormat(uint? numberFormatId, string? formatCode)
        {
            if (_numberingFormat.HasValue) throw new ApplicationException($"Numbering format style has alread been set");
            _numberingFormat = _excel.CreateNumberingFormat(numberFormatId, formatCode);
            return this;
        }

        public CellStyleBuilder SetPivotButton(bool value)
        {
            if (_pivotButton.HasValue) throw new ApplicationException($"Numbering format style has alread been set");
            _pivotButton = value;
            return this;
        }

        public CellStyleBuilder SetProtection(bool hidden, bool locked)
        {
            if (_protection != null) throw new ApplicationException($"Protection style has alread been set");
            _protection = _excel.GetProtection(hidden, locked);
            return this;
        }

        public CellStyleBuilder SetQuotePrefix(bool value)
        {
            if (_quotePrefix.HasValue) throw new ApplicationException($"Numbering format style has alread been set");
            _quotePrefix = value;
            return this;
        }
    }
}