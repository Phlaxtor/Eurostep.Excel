using DocumentFormat.OpenXml.Packaging;

namespace Eurostep.Excel;

public class ExcelReader
{
    private readonly SpreadsheetDocument? _spreasheet;

    public ExcelReader(Stream stream)
    {
        OpenSettings settings = new OpenSettings()
        {
            AutoSave = true,
        };
        _spreasheet = SpreadsheetDocument.Open(stream, false, settings);
    }
}