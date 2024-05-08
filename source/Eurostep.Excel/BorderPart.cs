using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    public readonly struct BorderPart
    {
        private readonly GeneralColor _color;
        private readonly BorderStyleValues _style;

        public BorderPart(BorderStyleValues style, GeneralColor color)
        {
            _style = style;
            _color = color;
        }

        public GeneralColor Color => _color;
        public BorderStyleValues Style => _style;

        public T ToBorder<T>() where T : BorderPropertiesType, new()
        {
            return new T
            {
                Style = _style,
                Color = _color.ToSpreadsheetColor<Color>()
            };
        }
    }
}