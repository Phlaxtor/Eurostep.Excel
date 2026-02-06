using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using FMA.Digimat.Toolbox.Excel.Model;

using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace FMA.Digimat.Toolbox.Excel
{
    public static partial class ExcelGenerationHelperMethods
    {
        public static uint? CreateBorder(this Stylesheet self, ExcelBorderStyle? style)
        {
            if (style is null)
            {
                return null;
            }
            return self.CreateBorder(
                style.BorderStyle,
                style.BorderColor
                );
        }

        public static uint? CreateBorder(
            this Stylesheet self,
            BorderStyleValues style,
            HexBinaryValue argbColor)
        {
            var border = new Border();
            border.TopBorder = new TopBorder
            {
                Style = new EnumValue<BorderStyleValues>(style),
                Color = new Spreadsheet.Color { Rgb = argbColor }
            };
            border.RightBorder = new RightBorder
            {
                Style = new EnumValue<BorderStyleValues>(style),
                Color = new Spreadsheet.Color { Rgb = argbColor }
            };
            border.BottomBorder = new BottomBorder
            {
                Style = new EnumValue<BorderStyleValues>(style),
                Color = new Spreadsheet.Color { Rgb = argbColor }
            };
            border.LeftBorder = new LeftBorder
            {
                Style = new EnumValue<BorderStyleValues>(style),
                Color = new Spreadsheet.Color { Rgb = argbColor }
            };
            self.Borders.Append(border);
            if (self.Borders.Count is null)
            {
                self.Borders.Count = 0;
            }
            else
            {
                self.Borders.Count++;
            }
            self.Save();
            return self.Borders.Count;
        }

        public static uint? CreateCellFormat(
            this Stylesheet self,
            uint? numberFormatId,
            uint? formatId,
            Alignment alignment,
            uint? fontIndex,
            uint? borderId,
            uint? fillIndex,
            Protection protection,
            bool? pivotButton,
            bool? quotePrefix)
        {
            if (self == null)
            {
                throw new ArgumentNullException("self",
                    "The provided Stylesheet in the extension method must not be null.");
            }

            if (self.CellFormats == null)
            {
                throw new ApplicationException("Stylesheet.CellFormats must not be null.");
            }

            var cellFormat = new CellFormat();

            if (pivotButton.HasValue)
            {
                cellFormat.PivotButton = BooleanValue.FromBoolean(pivotButton.Value);
            }

            if (quotePrefix.HasValue)
            {
                cellFormat.QuotePrefix = BooleanValue.FromBoolean(quotePrefix.Value);
            }

            if (protection != null)
            {
                cellFormat.Protection = protection;
                cellFormat.ApplyProtection = BooleanValue.FromBoolean(true);
            }

            if (formatId.HasValue)
            {
                cellFormat.FormatId = UInt32Value.FromUInt32(formatId.Value);
            }

            if (alignment != null)
            {
                cellFormat.Alignment = alignment;
                cellFormat.ApplyAlignment = BooleanValue.FromBoolean(true);
            }

            if (borderId.HasValue)
            {
                cellFormat.BorderId = UInt32Value.FromUInt32(borderId.Value);
                cellFormat.ApplyBorder = BooleanValue.FromBoolean(true);
            }

            if (fontIndex.HasValue)
            {
                cellFormat.FontId = UInt32Value.FromUInt32(fontIndex.Value);
                cellFormat.ApplyFont = BooleanValue.FromBoolean(true);
            }

            if (fillIndex.HasValue)
            {
                cellFormat.FillId = UInt32Value.FromUInt32(fillIndex.Value);
                cellFormat.ApplyFill = BooleanValue.FromBoolean(true);
            }

            if (numberFormatId.HasValue)
            {
                cellFormat.NumberFormatId = UInt32Value.FromUInt32(numberFormatId.Value);
                cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);

                //0 General
                //1 0
                //2 0.00
                //3 #,##0
                //4 #,##0.00
                //9 0%
                //10 0.00%
                //11 0.00E+00
                //12 # ?/?
                //13 # ??/??
                //14 mm-dd-yy
                //15 d-mmm-yy
                //16 d-mmm
                //17 mmm-yy
                //18 h:mm AM/PM
                //19 h:mm:ss AM/PM
                //20 h:mm
                //21 h:mm:ss
                //22 m/d/yy h:mm
                //37 #,##0 ;(#,##0)
                //38 #,##0 ;[Red](#,##0)
                //39 #,##0.00;(#,##0.00)
                //40 #,##0.00;[Red](#,##0.00)
                //45 mm:ss
                //46 [h]:mm:ss
                //47 mmss.0
                //48 ##0.0E+0
                //49 @
            }

            self.CellFormats.Append(cellFormat);
            if (self.CellFormats.Count is null)
            {
                self.CellFormats.Count = 0;
            }
            else
            {
                self.CellFormats.Count++;
            }
            self.Save();
            return self.CellFormats.Count;
        }

        public static uint CreateCustomNumberFormat(
            this Stylesheet self,
            string formatCode)
        {
            const uint lowestIdForCustomFormat = 164; // lower values are reserved for built-in formats
            if (self.NumberingFormats == null)
            {
                self.NumberingFormats = new();
                self.NumberingFormats.Count = 0;
            }
            uint numberFormatId = lowestIdForCustomFormat + self.NumberingFormats.Count;
            var numberFormat = new NumberingFormat() { FormatCode = formatCode, NumberFormatId = numberFormatId };
            self.NumberingFormats.Append(numberFormat);
            self.NumberingFormats.Count++;
            self.Save();
            return numberFormatId;
        }

        public static uint? CreateFill(this Stylesheet self, ExcelFillStyle? style)
        {
            if (style is null)
            {
                return null;
            }
            return self.CreateFill(
                style.PatternType,
                style.ForegroundColor,
                style.BackgroundColor
                );
        }

        public static uint? CreateFill(this Stylesheet self, PatternValues patternType, string? foregroundColor, string? backgroundColor)
        {
            if (self == null)
            {
                throw new ArgumentNullException("self",
                    "The provided Stylesheet in the extension method must not be null.");
            }

            if (self.Fills == null)
            {
                throw new ApplicationException("Stylesheet.Fills must not be null.");
            }

            var fill = new Fill();

            fill.PatternFill = new PatternFill { PatternType = patternType };

            if (string.IsNullOrEmpty(foregroundColor) == false)
            {
                fill.PatternFill.Append(new ForegroundColor { Rgb = new HexBinaryValue(foregroundColor) });
            }

            if (string.IsNullOrEmpty(backgroundColor) == false)
            {
                fill.PatternFill.Append(new BackgroundColor { Rgb = new HexBinaryValue(backgroundColor) });
            }

            self.Fills.Append(fill);
            if (self.Fills.Count is null)
            {
                self.Fills.Count = 0;
            }
            else
            {
                self.Fills.Count++;
            }
            self.Save();
            return self.Fills.Count;
        }

        public static uint? CreateFont(this Stylesheet self, ExcelFontStyle? style)
        {
            if (style is null)
            {
                return null;
            }
            return self.CreateFont(
                style.FontName,
                style.FontSize,
                style.IsBold,
                style.IsItalic,
                style.UnderlineType,
                style.FontColor,
                style.VerticalAlignment,
                style.IsStrike,
                style.IsCondense,
                style.IsExtend,
                style.HasShadow
                );
        }

        public static uint? CreateFont(
            this Stylesheet self,
            string fontName,
            double? fontSize,
            bool? isBold,
            bool? isItalic,
            UnderlineValues? underlineType,
            HexBinaryValue argbColor,
            VerticalAlignmentRunValues? verticalAlignment,
            bool? isStrike,
            bool? isCondense,
            bool? isExtend,
            bool? hasShadow)
        {
            if (self == null)
            {
                throw new ArgumentNullException("self",
                    "The provided Stylesheet in the extension method must not be null.");
            }

            if (self.Fonts == null)
            {
                throw new ApplicationException("Stylesheet.Fonts must not be null.");
            }

            var font = new Font();
            if (!string.IsNullOrEmpty(fontName))
            {
                font.FontName = new FontName { Val = StringValue.ToString(fontName) };
            }

            if (fontSize.HasValue)
            {
                font.FontSize = new FontSize { Val = DoubleValue.ToDouble(fontSize.Value) };
            }

            if (isBold.HasValue)
            {
                font.Bold = new Bold { Val = BooleanValue.ToBoolean(isBold.Value) };
            }

            if (isItalic.HasValue)
            {
                font.Italic = new Italic { Val = BooleanValue.ToBoolean(isItalic.Value) };
            }

            if (underlineType.HasValue)
            {
                font.Underline = new Underline { Val = new EnumValue<UnderlineValues>(underlineType.Value) };
            }

            if (isStrike.HasValue)
            {
                font.Strike = new Strike { Val = BooleanValue.ToBoolean(isStrike.Value) };
            }

            if (isCondense.HasValue)
            {
                font.Condense = new Condense { Val = BooleanValue.ToBoolean(isCondense.Value) };
            }

            if (isExtend.HasValue)
            {
                font.Extend = new Extend { Val = BooleanValue.ToBoolean(isExtend.Value) };
            }

            if (hasShadow.HasValue)
            {
                font.Shadow = new Shadow { Val = BooleanValue.ToBoolean(hasShadow.Value) };
            }

            if (verticalAlignment.HasValue)
            {
                font.VerticalTextAlignment = new VerticalTextAlignment
                {
                    Val = new EnumValue<VerticalAlignmentRunValues>(verticalAlignment.Value)
                };
            }

            font.Color = new Spreadsheet.Color { Rgb = argbColor };

            self.Fonts.Append(font);
            if (self.Fonts.Count is null)
            {
                self.Fonts.Count = 0;
            }
            else
            {
                self.Fonts.Count++;
            }
            self.Save();
            return self.Fonts.Count;
        }

        public static uint? CreateSolidFill(this Stylesheet self, HexBinaryValue argbColor)
        {
            if (self == null)
            {
                throw new ArgumentNullException("self",
                    "The provided Stylesheet in the extension method must not be null.");
            }

            if (self.Fills == null)
            {
                throw new ApplicationException("Stylesheet.Fills must not be null.");
            }

            var fill = new Fill();

            fill.PatternFill = new PatternFill { PatternType = PatternValues.Solid };
            var foregroundColor = new ForegroundColor { Rgb = argbColor };
            fill.PatternFill.Append(foregroundColor);
            fill.PatternFill.Append(new BackgroundColor { Indexed = (UInt32Value)64U });

            self.Fills.Append(fill);
            if (self.Fills.Count is null)
            {
                self.Fills.Count = 0;
            }
            else
            {
                self.Fills.Count++;
            }
            self.Save();
            return self.Fills.Count;
        }
    }
}