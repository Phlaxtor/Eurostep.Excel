namespace Eurostep.Excel
{
    public static class ExcelFile
    {
        public static IExcelDocument GetExcelDocument(Stream stream)
        {
            return new ExcelDocument(stream);
        }

        public static IExcelDocument GetExcelDocument(string path)
        {
            return new ExcelDocument(path);
        }

        public static ISheet GetExcelSheet(Stream stream, string sheetName)
        {
            ExcelDocument document = new ExcelDocument(stream, false);
            ISheet? sheet = document.GetSheetByName(sheetName);
            return sheet ?? throw new ArgumentException($"Can not find sheet with name '{sheetName}'", nameof(sheetName));
        }

        public static ISheet GetExcelSheet(string path, string sheetName)
        {
            ExcelDocument document = new ExcelDocument(path, false);
            ISheet? sheet = document.GetSheetByName(sheetName);
            return sheet ?? throw new ArgumentException($"Can not find sheet with name '{sheetName}'", nameof(sheetName));
        }

        public static ISheet GetExcelSheet(Stream stream, int index)
        {
            ExcelDocument document = new ExcelDocument(stream, false);
            ISheet? sheet = document.GetSheetByPosition(index);
            return sheet ?? throw new ArgumentException($"Can not find sheet at index '{index}'", nameof(index));
        }

        public static ISheet GetExcelSheet(string path, int index)
        {
            ExcelDocument document = new ExcelDocument(path, false);
            ISheet? sheet = document.GetSheetByPosition(index);
            return sheet ?? throw new ArgumentException($"Can not find sheet at index '{index}'", nameof(index));
        }

        public static ISheetReader GetSheetReader(Stream stream, string sheetName, params string[] headers)
        {
            ExcelDocument document = new ExcelDocument(stream, false);
            ISheet? sheet = document.GetSheetByName(sheetName);
            _ = sheet ?? throw new ArgumentException($"Can not find sheet with name '{sheetName}'", nameof(sheetName));
            ISheetReader reader = sheet.GetReader();
            if (headers.Length > 0)
            {
                if (reader.ReadToHeaders(headers) == false) throw new ApplicationException($"Headers not found in sheet '{sheet.Name}'");
            }
            return reader;
        }

        public static ISheetReader GetSheetReader(string path, string sheetName, params string[] headers)
        {
            ExcelDocument document = new ExcelDocument(path, false);
            ISheet? sheet = document.GetSheetByName(sheetName);
            _ = sheet ?? throw new ArgumentException($"Can not find sheet with name '{sheetName}'", nameof(sheetName));
            ISheetReader reader = sheet.GetReader();
            if (headers.Length > 0)
            {
                if (reader.ReadToHeaders(headers) == false) throw new ApplicationException($"Headers not found in sheet '{sheet.Name}'");
            }
            return reader;
        }

        public static ISheetReader GetSheetReader(Stream stream, int index, params string[] headers)
        {
            ExcelDocument document = new ExcelDocument(stream, false);
            ISheet? sheet = document.GetSheetByPosition(index);
            _ = sheet ?? throw new ArgumentException($"Can not find sheet at index '{index}'", nameof(index));
            ISheetReader reader = sheet.GetReader();
            if (headers.Length > 0)
            {
                if (reader.ReadToHeaders(headers) == false) throw new ApplicationException($"Headers not found in sheet '{sheet.Name}'");
            }
            return reader;
        }

        public static ISheetReader GetSheetReader(string path, int index, params string[] headers)
        {
            ExcelDocument document = new ExcelDocument(path, false);
            ISheet? sheet = document.GetSheetByPosition(index);
            _ = sheet ?? throw new ArgumentException($"Can not find sheet at index '{index}'", nameof(index));
            ISheetReader reader = sheet.GetReader();
            if (headers.Length > 0)
            {
                if (reader.ReadToHeaders(headers) == false) throw new ApplicationException($"Headers not found in sheet '{sheet.Name}'");
            }
            return reader;
        }
    }
}