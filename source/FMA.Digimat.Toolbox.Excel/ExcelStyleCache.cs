using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FMA.Digimat.Toolbox.Excel.Model;

namespace FMA.Digimat.Toolbox.Excel
{
    public interface IExcelStyleCache
    {
        uint? GetBorderStyle<T>() where T : ExcelBorderStyle, new();

        uint? GetBorderStyle(ExcelBorderStyle? style);

        uint? GetCellStyle<T>() where T : ExcelCellStyle, new();

        uint? GetCellStyle(ExcelCellStyle? style);

        uint? GetFillStyle<T>() where T : ExcelFillStyle, new();

        uint? GetFillStyle(ExcelFillStyle? style);

        uint? GetFontStyle<T>() where T : ExcelFontStyle, new();

        uint? GetFontStyle(ExcelFontStyle? style);

        uint? GetNumberingFormatStyle<T>() where T : ExcelNumberingFormatStyle, new();

        uint? GetNumberingFormatStyle(ExcelNumberingFormatStyle? style);
    }

    public sealed class ExcelStyleCache : IExcelStyleCache
    {
        private readonly Dictionary<Type, uint?> _borderStyles = [];
        private readonly Dictionary<Type, uint?> _cellStyles = [];
        private readonly Dictionary<Type, uint?> _fillStyles = [];
        private readonly Dictionary<Type, uint?> _fontStyles = [];
        private readonly Stylesheet _stylesheet;

        public ExcelStyleCache(SpreadsheetDocument excel)
        {
            _stylesheet = excel.GetStylesheet();
            GetCellStyle(new DefaultExcelCellStyle());
            _stylesheet.CreateFill(PatternValues.Gray125, null, null);
        }

        public Stylesheet Stylesheet => _stylesheet;

        public uint? GetBorderStyle(ExcelBorderStyle? style)
        {
            if (style is null)
            {
                return null;
            }

            if (_borderStyles.TryGetValue(style.GetType(), out uint? result) == false)
            {
                _borderStyles[style.GetType()] = result = _stylesheet.CreateBorder(style);
            }

            return result;
        }

        public uint? GetBorderStyle<T>() where T : ExcelBorderStyle, new()
        {
            return GetBorderStyle(new T());
        }

        public uint? GetCellStyle<T>() where T : ExcelCellStyle, new()
        {
            return GetCellStyle(new T());
        }

        public uint? GetCellStyle(ExcelCellStyle? style)
        {
            if (style is null)
            {
                return null;
            }

            if (_cellStyles.TryGetValue(style.GetType(), out uint? result))
            {
                return result;
            }

            uint? formatId = null;
            uint? borderStyle = GetBorderStyle(style.BorderStyle);
            uint? fillStyle = GetFillStyle(style.FillStyle);
            uint? fontStyle = GetFontStyle(style.FontStyle);
            uint? numberFormatId = GetNumberingFormatStyle(style.NumberingFormatStyle);
            _cellStyles[style.GetType()] = result = _stylesheet.CreateCellFormat(numberFormatId, formatId, style.Alignment, fontStyle, borderStyle, fillStyle, style.Protection, style.PivotButton, style.QuotePrefix);
            return result;
        }

        public uint? GetFillStyle<T>() where T : ExcelFillStyle, new()
        {
            return GetFillStyle(new T());
        }

        public uint? GetFillStyle(ExcelFillStyle? style)
        {
            if (style is null)
            {
                return null;
            }

            if (_fillStyles.TryGetValue(style.GetType(), out uint? result) == false)
            {
                _fillStyles[style.GetType()] = result = _stylesheet.CreateFill(style);
            }

            return result;
        }

        public uint? GetFontStyle<T>() where T : ExcelFontStyle, new()
        {
            return GetFontStyle(new T());
        }

        public uint? GetFontStyle(ExcelFontStyle? style)
        {
            if (style is null)
            {
                return null;
            }

            if (_fontStyles.TryGetValue(style.GetType(), out uint? result) == false)
            {
                _fontStyles[style.GetType()] = result = _stylesheet.CreateFont(style);
            }

            return result;
        }

        public uint? GetNumberingFormatStyle<T>() where T : ExcelNumberingFormatStyle, new()
        {
            return GetNumberingFormatStyle(new T());
        }

        public uint? GetNumberingFormatStyle(ExcelNumberingFormatStyle? style)
        {
            if (style is null)
            {
                return null;
            }
            // TODO: Do this in a proper way, not a prio since it is not used
            return style.NumberFormatId;
        }
    }
}