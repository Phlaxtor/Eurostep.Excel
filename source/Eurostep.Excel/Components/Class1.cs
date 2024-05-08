using DocumentFormat.OpenXml.Spreadsheet;
using System.Drawing;

namespace Eurostep.Excel.Components
{
    internal class TestWriter
    {
        private readonly ExcelWriter _writer;
        private uint? _descriptionFixedStyle;
        private uint? _descriptionGeneralStyle;
        private uint? _descriptionMandatoryStyle;
        private uint? _descriptionOptionalStyle;
        private uint? _fixedHeaderStyle;
        private uint? _itemIdCellStyle;
        private uint? _mandatoryHeaderStyle;
        private uint? _optionalHeaderStyle;
        private uint? _versionCellStyle;

        public TestWriter(ExcelWriter writer)
        {
            _writer = writer;
        }

        public uint GetGeneralStyle(GeneralColor cellFill, GeneralColor fontColor, double fontSize = 11)
        {
            var allignment = _writer.GetAlignment(HorizontalAlignmentValues.Left, VerticalAlignmentValues.Center);
            var fontId = _writer.CreateFont("Calibri", fontSize, false, false, UnderlineValues.None, fontColor);
            var fillId = _writer.CreateFill(PatternValues.Solid, cellFill);
            var protection = _writer.GetProtection(false, false);
            var cellFormat = _writer.CreateCellFormat(null, null, allignment, fontId, null, fillId, protection);
            _writer.Spreasheet.Save();
            return cellFormat;
        }

        public uint GetHeaderStyle(GeneralColor cellFill, GeneralColor fontColor, double fontSize = 11)
        {
            var borderStyle = new BorderPart(BorderStyleValues.Thin, KnownColor.DarkGray);
            var allignment = _writer.GetAlignment(HorizontalAlignmentValues.Left, VerticalAlignmentValues.Center);
            var fontId = _writer.CreateFont("Calibri", fontSize, true, false, UnderlineValues.None, fontColor);
            var borderId = _writer.CreateBorder(top: borderStyle, bottom: borderStyle);
            var fillId = _writer.CreateFill(PatternValues.Solid, cellFill);
            var protection = _writer.GetProtection(false, true);
            var cellFormat = _writer.CreateCellFormat(null, null, allignment, fontId, borderId, fillId, protection);
            _writer.Spreasheet.Save();
            return cellFormat;
        }

        public uint GetItemIdCellStyle()
        {
            if (_itemIdCellStyle.HasValue) return _itemIdCellStyle.Value;
            _itemIdCellStyle = GetNumberCellStyle(164, "00000000", KnownColor.Black);
            return _itemIdCellStyle.Value;
        }

        public uint GetLightBlueGeneralStyle()
        {
            //LightBlue
            if (_descriptionFixedStyle.HasValue) return _descriptionFixedStyle.Value;
            _descriptionFixedStyle = GetGeneralStyle(new GeneralColor("b4c6e7"), KnownColor.Black);
            return _descriptionFixedStyle.Value;
        }

        public uint GetLightBlueHeaderStyle()
        {
            //LightBlue
            if (_fixedHeaderStyle.HasValue) return _fixedHeaderStyle.Value;
            _fixedHeaderStyle = GetHeaderStyle(new GeneralColor("b4c6e7"), KnownColor.Black);
            return _fixedHeaderStyle.Value;
        }

        public uint GetLightGrayGeneralStyle()
        {
            //LightGray
            if (_descriptionGeneralStyle.HasValue) return _descriptionGeneralStyle.Value;
            _descriptionGeneralStyle = GetGeneralStyle(new GeneralColor("ededed"), KnownColor.Black);
            return _descriptionGeneralStyle.Value;
        }

        public uint GetLightGreenGeneralStyle()
        {
            //LightGreen
            if (_descriptionMandatoryStyle.HasValue) return _descriptionMandatoryStyle.Value;
            _descriptionMandatoryStyle = GetGeneralStyle(new GeneralColor("c6e0b4"), KnownColor.Black);
            return _descriptionMandatoryStyle.Value;
        }

        public uint GetLightGreenHeaderStyle()
        {
            //LightGreen
            if (_mandatoryHeaderStyle.HasValue) return _mandatoryHeaderStyle.Value;
            _mandatoryHeaderStyle = GetHeaderStyle(new GeneralColor("c6e0b4"), KnownColor.Black);
            return _mandatoryHeaderStyle.Value;
        }

        public uint GetLightOrangeGeneralStyle()
        {
            //LightOrange
            if (_descriptionOptionalStyle.HasValue) return _descriptionOptionalStyle.Value;
            _descriptionOptionalStyle = GetGeneralStyle(new GeneralColor("f8cbad"), KnownColor.Black);
            return _descriptionOptionalStyle.Value;
        }

        public uint GetLightOrangeHeaderStyle()
        {
            //LightOrange
            if (_optionalHeaderStyle.HasValue) return _optionalHeaderStyle.Value;
            _optionalHeaderStyle = GetHeaderStyle(new GeneralColor("f8cbad"), KnownColor.Black);
            return _optionalHeaderStyle.Value;
        }

        public uint GetNumberCellStyle(uint numFmtId, string formatCode, GeneralColor fontColor, double fontSize = 11)
        {
            uint numId = _writer.CreateNumberingFormat(numFmtId, formatCode);
            Alignment allignment = _writer.GetAlignment(HorizontalAlignmentValues.Left, VerticalAlignmentValues.Center);
            uint fontId = _writer.CreateFont("Calibri", fontSize, false, false, UnderlineValues.None, fontColor);
            Protection protection = _writer.GetProtection(false, false);
            uint cellFormat = _writer.CreateCellFormat(numFmtId, null, allignment, fontId, null, null, protection);
            _writer.Spreasheet.Save();
            return cellFormat;
        }

        public uint GetVersionCellStyle()
        {
            if (_versionCellStyle.HasValue) return _versionCellStyle.Value;
            _versionCellStyle = GetNumberCellStyle(165, "000", KnownColor.Black);
            return _versionCellStyle.Value;
        }
    }
}