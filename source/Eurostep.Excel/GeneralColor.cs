using DocumentFormat.OpenXml;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;

namespace Eurostep.Excel
{
    public readonly struct GeneralColor
    {
        public static readonly GeneralColor Empty = new GeneralColor(Color.Empty);
        private readonly Color _color;
        private readonly double? _tint;

        public GeneralColor(Color color, double? tint = null)
        {
            _color = color;
            _tint = tint;
        }

        public GeneralColor(KnownColor color, double? tint = null)
        {
            _color = Color.FromKnownColor(color);
            _tint = tint;
        }

        public GeneralColor(string hex, double? tint = null)
        {
            int color = Convert.ToInt32(hex, 16);
            _color = Color.FromArgb(color);
            _tint = tint;
        }

        public GeneralColor(byte red, byte green, byte blue, byte? alpha = null, double? tint = null)
        {
            if (alpha.HasValue)
            {
                _color = Color.FromArgb(alpha.Value, red, green, blue);
            }
            else
            {
                _color = Color.FromArgb(red, green, blue);
            }
            _tint = tint;
        }

        public string Hex => _color.R.ToString("X2") + _color.G.ToString("X2") + _color.B.ToString("X2");
        public Color SystemDrawingColor => _color;

        public static implicit operator Color(GeneralColor c)
        {
            return c.SystemDrawingColor;
        }

        public static implicit operator GeneralColor(Color c)
        {
            return new GeneralColor(c);
        }

        public static implicit operator GeneralColor(KnownColor c)
        {
            return new GeneralColor(c);
        }

        public override bool Equals([NotNullWhen(true)] object? obj)
        {
            if (obj is GeneralColor gc) return _color.Equals(gc._color);
            if (obj is Color c) return _color.Equals(c);
            return false;
        }

        public override int GetHashCode()
        {
            return _color.GetHashCode();
        }

        public T ToSpreadsheetColor<T>(bool? auto = default)
            where T : DocumentFormat.OpenXml.Spreadsheet.ColorType, new()
        {
            return new T
            {
                Rgb = new HexBinaryValue { Value = Hex },
                Auto = auto,
                Tint = _tint
            };
        }

        public override string ToString()
        {
            return Hex;
        }
    }
}