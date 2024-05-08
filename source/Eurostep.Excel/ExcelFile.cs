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
            var document = new ExcelDocument(stream, false);
            var sheet = document.GetSheetByName(sheetName);
            return sheet ?? throw new ArgumentException($"Can not find sheet with name '{sheetName}'", nameof(sheetName));
        }

        public static ISheet GetExcelSheet(string path, string sheetName)
        {
            var document = new ExcelDocument(path, false);
            var sheet = document.GetSheetByName(sheetName);
            return sheet ?? throw new ArgumentException($"Can not find sheet with name '{sheetName}'", nameof(sheetName));
        }

        public static ISheet GetExcelSheet(Stream stream, int index)
        {
            var document = new ExcelDocument(stream, false);
            var sheet = document.GetSheetByPosition(index);
            return sheet ?? throw new ArgumentException($"Can not find sheet at index '{index}'", nameof(index));
        }

        public static ISheet GetExcelSheet(string path, int index)
        {
            var document = new ExcelDocument(path, false);
            var sheet = document.GetSheetByPosition(index);
            return sheet ?? throw new ArgumentException($"Can not find sheet at index '{index}'", nameof(index));
        }

        public static ISheetReader GetSheetReader(Stream stream, string sheetName, params string[] headers)
        {
            var document = new ExcelDocument(stream, false);
            var sheet = document.GetSheetByName(sheetName);
            _ = sheet ?? throw new ArgumentException($"Can not find sheet with name '{sheetName}'", nameof(sheetName));
            var reader = sheet.GetReader();
            if (headers.Length > 0)
            {
                if (reader.ReadToHeaders(headers) == false) throw new ApplicationException($"Headers not found in sheet '{sheet.Name}'");
            }
            return reader;
        }

        public static ISheetReader GetSheetReader(string path, string sheetName, params string[] headers)
        {
            var document = new ExcelDocument(path, false);
            var sheet = document.GetSheetByName(sheetName);
            _ = sheet ?? throw new ArgumentException($"Can not find sheet with name '{sheetName}'", nameof(sheetName));
            var reader = sheet.GetReader();
            if (headers.Length > 0)
            {
                if (reader.ReadToHeaders(headers) == false) throw new ApplicationException($"Headers not found in sheet '{sheet.Name}'");
            }
            return reader;
        }

        public static ISheetReader GetSheetReader(Stream stream, int index, params string[] headers)
        {
            var document = new ExcelDocument(stream, false);
            var sheet = document.GetSheetByPosition(index);
            _ = sheet ?? throw new ArgumentException($"Can not find sheet at index '{index}'", nameof(index));
            var reader = sheet.GetReader();
            if (headers.Length > 0)
            {
                if (reader.ReadToHeaders(headers) == false) throw new ApplicationException($"Headers not found in sheet '{sheet.Name}'");
            }
            return reader;
        }

        public static ISheetReader GetSheetReader(string path, int index, params string[] headers)
        {
            var document = new ExcelDocument(path, false);
            var sheet = document.GetSheetByPosition(index);
            _ = sheet ?? throw new ArgumentException($"Can not find sheet at index '{index}'", nameof(index));
            var reader = sheet.GetReader();
            if (headers.Length > 0)
            {
                if (reader.ReadToHeaders(headers) == false) throw new ApplicationException($"Headers not found in sheet '{sheet.Name}'");
            }
            return reader;
        }
    }
}