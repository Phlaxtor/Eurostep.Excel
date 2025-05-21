using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    public readonly struct BorderPart
    {
        public static readonly BorderPart Empty = new BorderPart();
        private readonly GeneralColor _color;
        private readonly BorderStyleValues _style;

        public BorderPart()
        {
            _color = GeneralColor.Empty;
            _style = BorderStyleValues.None;
        }

        public BorderPart(BorderStyleValues style, GeneralColor color)
        {
            _color = color;
            _style = style;
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