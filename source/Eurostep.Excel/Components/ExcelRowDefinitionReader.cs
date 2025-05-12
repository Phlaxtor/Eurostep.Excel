using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Eurostep.Excel;
using System.Globalization;
using System.Reflection;

namespace EPPSPrio.Components.Excel;

public class ExcelRowDefinitionReader
{
    internal static Log LogError;

    public delegate void Log(string message);

    private bool CheckColumnNames<T>(T rowInstance, Dictionary<string, string> headerRow)
        where T : ExcelRowDefinition
    {
        rowInstance.HeadingsWithColumnNames = [];
        bool ret = true;
        Type objtype = rowInstance.GetType();
        // Loop through all properties
        List<string> fieldNames = [];
        foreach (PropertyInfo p in objtype.GetProperties())
        {
            ExcelColumnAttribute? fieldNameAttribute = p.GetCustomAttribute<ExcelColumnAttribute>(false);
            if (fieldNameAttribute != null)
            {
                string name = fieldNameAttribute.Heading;
                KeyValuePair<string, string> cell = headerRow.FirstOrDefault(p => name.Equals(p.Value.Trim(), StringComparison.InvariantCultureIgnoreCase));
                if (cell.Key == null)
                {
                    LogError($"Sheet {rowInstance.SheetName}: Heading \"{name}\" not found in header row {headerRow}.");
                    ret = false;
                }
                rowInstance.HeadingsWithColumnNames[name] = cell.Key;
            }
        }

        return ret;
    }

    private bool CheckSheet<T>(T rowInstance, Dictionary<uint, Dictionary<string, string>> sheetArea)
        where T : ExcelRowDefinition
    {
        if (sheetArea.Count() < rowInstance.FirstDataRow)
        {
            LogError($"Sheet {rowInstance.SheetName} contains no data.");
            return false;
        }

        if (!sheetArea.ContainsKey(rowInstance.HeaderRow) || !CheckColumnNames(rowInstance, sheetArea[rowInstance.HeaderRow]))
        {
            LogError($"Sheet {rowInstance.SheetName} has incorrect header (row {rowInstance.HeaderRow}).");
            return false;
        }
        return true;
    }

    private string GetCellValue(SpreadsheetDocument self, Cell cell, SharedStringItem[] sharedStrings)
    {
        string foundValue = string.Empty;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            foundValue = sharedStrings[int.Parse(cell.CellValue!.Text)].InnerText;
        }
        else
        {
            if (cell.CellValue != null)
            {
                foundValue = cell.CellValue.Text;
            }
        }

        return foundValue;
    }

    private DateTime? GetDateTimeValue(Dictionary<string, string> row, string key)
    {
        DateTime? parsedDate = null;
        string? value = GetStringValue(row, key);
        if (!string.IsNullOrWhiteSpace(value))
        {
            try
            {
                double d = double.Parse(value);
                parsedDate = DateTime.FromOADate(d);
            }
            catch (FormatException) // catching the exception for those extreme rare occasions this might happen, don't want it in the flow since it will slow down the process.
            {
                if (value.Contains('-'))
                {
                    parsedDate = DateTime.ParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                }
            }
        }
        return parsedDate;
    }

    private int? GetIntValue(Dictionary<string, string> row, string key)
    {
        int? ret = null;
        string? value = GetStringValue(row, key);
        if (value.IsInteger())
        {
            ret = int.Parse(value);
        }
        return ret;
    }

    private Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetArea(WorksheetPart self, Cell upperRightCell, Cell lowerLeftCell)
    {
        if (upperRightCell == null)
        {
            throw new ArgumentNullException("upperRightCell", "The provided Cell must not be null.");
        }

        if (lowerLeftCell == null)
        {
            throw new ArgumentNullException("lowerLeftCell", "The provided Cell must not be null.");
        }

        string columnStart = upperRightCell.GetColumnName();
        string columnEnd = lowerLeftCell.GetColumnName();
        uint rowStart = upperRightCell.GetRowIndex();
        uint rowEnd = lowerLeftCell.GetRowIndex();

        return GetRowsExcelSheetArea(self, columnStart, rowStart, columnEnd, rowEnd);
    }

    private Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetArea(WorksheetPart self, SheetDimension area)
    {
        if (area == null)
        {
            throw new ArgumentNullException("area", "The provided SheetDimension must not be null.");
        }

        if (!area.Reference.HasValue)
        {
            throw new ArgumentException("The provided SheetDimension.Reference must have an value.", "area");
        }

        string[] startEndValues = area.Reference.Value.Split(':');
        string? startValue = startEndValues.FirstOrDefault();
        string? endValue = startEndValues.LastOrDefault();
        string columnStart = ExcelUtilityMethods.GetColumnName(startValue);
        string columnEnd = ExcelUtilityMethods.GetColumnName(endValue);
        uint rowStart = ExcelUtilityMethods.GetRowIndex(startValue);
        uint rowEnd = ExcelUtilityMethods.GetRowIndex(endValue);

        return GetRowsExcelSheetArea(self, columnStart, rowStart, columnEnd, rowEnd);
    }

    private Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetArea(WorksheetPart self, string columnStart, uint rowStart, string columnEnd, uint rowEnd)
    {
        Dictionary<uint, Dictionary<string, string>> returnArrayOfRows = [];
        Dictionary<string, string> indexedRow = [];
        IEnumerable<Cell> cells =
            self.Worksheet.Descendants<Cell>().Where(
            c =>
                c.CellValue != null &&
                c.CompareColumn(columnStart) >= 0 &&
                c.CompareColumn(columnEnd) <= 0 &&
                c.GetRowIndex() >= rowStart && c.GetRowIndex() <= rowEnd)
                .OrderBy(q => q.GetRowIndex())
                .ThenBy(r => r.GetColumnIndex());

        SpreadsheetDocument spreadsheetDocument = self.OpenXmlPackage.GetSpreadsheetDocument();
        SharedStringItem[] sharedStringItems = Array.Empty<SharedStringItem>();
        IEnumerable<SharedStringTablePart> sharedStringPartCollection = spreadsheetDocument.WorkbookPart!.GetPartsOfType<SharedStringTablePart>();
        if (sharedStringPartCollection.IsNullOrEmpty())
        {
            // FIXME: properly supply an ILogger
            Console.Error.WriteLine("High performance Excel reading extensions only work if the file contains a SharedStringTablePart");
        }
        else
        {
            // there supposed to be exactly one shared string part
            SharedStringTablePart shareStringPart = sharedStringPartCollection.Single();
            sharedStringItems = shareStringPart!.SharedStringTable.Elements<SharedStringItem>().ToArray();
        }

        foreach (Cell cell in cells)
        {
            string columnName = cell.GetColumnName();
            uint rowNumber = cell.GetRowIndex();
            string cellValue = GetCellValue(spreadsheetDocument, cell, sharedStringItems);
            if (!returnArrayOfRows.TryGetValue(rowNumber, out Dictionary<string, string> rowInfo))
            {
                rowInfo = [];
                returnArrayOfRows[rowNumber] = rowInfo;
            }

            rowInfo[columnName] = cellValue;
        }
        return returnArrayOfRows;
    }

    private Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetArea(WorksheetPart self)
    {
        if (self != null)
        {
            if (self.Worksheet.SheetDimension != null)
            {
                return GetRowsExcelSheetArea(self, self.Worksheet.SheetDimension);
            }

            Cell? upperRightCell = self.Worksheet.LastChild?.FirstChild?.FirstChild as Cell;
            if (upperRightCell == null)
            {
                upperRightCell = self.Worksheet.Descendants<Row>()?.FirstOrDefault()?.Descendants<Cell>()?.FirstOrDefault();
            }
            Cell? lowerLeftCell = self.Worksheet.LastChild?.LastChild?.LastChild as Cell;
            if (lowerLeftCell == null)
            {
                lowerLeftCell = self.Worksheet.Descendants<Row>()?.LastOrDefault()?.Descendants<Cell>()?.LastOrDefault();
            }
            return GetRowsExcelSheetArea(self, upperRightCell, lowerLeftCell);
        }

        return null;
    }

    private string? GetStringValue(Dictionary<string, string> row, string key)
    {
        if (row.ContainsKey(key))
        {
            return (row[key]);
        }
        return null;
    }

    private WorksheetPart InitializeSpreadsheet<T>(T rowInstance, SpreadsheetDocument spreadsheet)
        where T : ExcelRowDefinition
    {
        spreadsheet.InitializeSpreadsheet(rowInstance.SheetName);
        WorksheetPart worksheet = spreadsheet.GetWorksheetPartBySheetName(rowInstance.SheetName);
        Type objtype = rowInstance.GetType();
        foreach (PropertyInfo p in objtype.GetProperties())
        {
            object? fieldNameAttribute = p.GetCustomAttributes(false).FirstOrDefault(z => z is ExcelColumnAttribute);
            if (fieldNameAttribute != null)
            {
                string key = ((ExcelColumnAttribute)fieldNameAttribute).Column;
                string heading = ((ExcelColumnAttribute)fieldNameAttribute).Heading;
                worksheet.SetColumnsData(key, 20);
                worksheet.WriteValueInCell(key, rowInstance.HeaderRow, heading);
                if (rowInstance.DescriptionRow.HasValue)
                {
                    string description = ((ExcelColumnAttribute)fieldNameAttribute).Description;
                    worksheet.WriteValueInCell(key, rowInstance.DescriptionRow.Value, description);
                }
            }
        }
        return worksheet;
    }

    private List<T> ReadData<T>(Stream stream, string sheetName, Log errorLogger)
        where T : ExcelRowDefinition, new()
    {
        LogError = errorLogger;
        List<T> data = [];
        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false))
        {
            T t = new T()
            {
                SheetName = sheetName
            };
            WorksheetPart worksheetPart = spreadsheet.GetWorksheetPartBySheetName(sheetName);
            ReadDataFromWorksheetPart(sheetName, data, t, worksheetPart);
        }
        return data;
    }

    /// <summary>
    /// Method that reads a named sheet if there are multiple or the first and only sheet.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="stream"></param>
    /// <param name="sheetName"></param>
    /// <param name="errorLogger"></param>
    /// <returns></returns>
    private List<T> ReadDataFromOnlySheetOrNamed<T>(Stream stream, string sheetName, Log errorLogger)
        where T : ExcelRowDefinition, new()
    {
        // TODO: Should use a standard logger and contrib to Eurostep.Implementation (ABI 2024-02)
        LogError = errorLogger;
        List<T> data = [];
        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false))
        {
            T t = new T()
            {
                SheetName = sheetName
            };
            // SAS Excel extensions count Sheet descendants instead of WorksheetParts
            //int? worksheets = spreadsheet.WorkbookPart?.WorksheetParts.Count();
            int? worksheets = spreadsheet.WorkbookPart?.Workbook.Descendants<Sheet>().Count();

            WorksheetPart worksheetPart = worksheets switch
            {
                // cannot proceed without at least one sheet
                0 or null => throw new ArgumentNullException(nameof(stream),
                    "The spreadsheet in the specified file stream does not contain any work sheets."),
                // get the first sheet if there is only one
                1 => spreadsheet.GetWorksheetPartByIndex(0),
                // get the sheet named {sheetName} if there are multiple
                _ => spreadsheet.GetWorksheetPartBySheetName(sheetName)
            };

            if (worksheetPart == null)
            {
                if (worksheets > 1)
                {
                    throw new ArgumentOutOfRangeException(
                            $"A work sheet with the name '{sheetName}' cannot be found.");
                }
                else
                {
                    // there is at least one Sheet but no corresponding WorksheetPart instance
                    throw new InvalidOperationException("Malformed XLSX file is suspected");
                }
            }

            ReadDataFromWorksheetPart(sheetName, data, t, worksheetPart);
        }
        return data;
    }

    private void ReadDataFromWorksheetPart<T>(string sheetName, List<T> data, T t, WorksheetPart worksheetPart)
        where T : ExcelRowDefinition, new()
    {
        if (worksheetPart == null)
        {
            LogError($"Sheet {sheetName} not found in the input file.");
            throw new InvalidDataException($"Invalid input file: Sheet {sheetName} not found in the input file.");
        }
        ;
        Dictionary<uint, Dictionary<string, string>> sheetArea = GetRowsExcelSheetArea(worksheetPart);

        if (!CheckSheet(t, sheetArea))
        {
            throw new InvalidDataException($"Invalid input file: Invalid data in sheet {sheetName}");
        }

        foreach (KeyValuePair<uint, Dictionary<string, string>> row in sheetArea)
        {
            if (row.Key < t.FirstDataRow) { continue; }
            T instance = new T
            {
                SheetName = sheetName,
                HeadingsWithColumnNames = t.HeadingsWithColumnNames
            };
            ReadRow(instance, row);
            data.Add(instance);
        }
    }

    private void ReadRow<T>(T rowInstance, KeyValuePair<uint, Dictionary<string, string>> rowData)
        where T : ExcelRowDefinition
    {
        rowInstance.RowId = rowData.Key;
        rowInstance.DetailsForLogging = $"sheet: {rowInstance.SheetName}; row: {rowInstance.RowId}";
        Type objtype = rowInstance.GetType();
        foreach (PropertyInfo p in objtype.GetProperties())
        {
            object? fieldNameAttribute = p.GetCustomAttributes(false).FirstOrDefault(z => z is ExcelColumnAttribute);
            if (fieldNameAttribute != null)
            {
                string key = rowInstance.HeadingsWithColumnNames[((ExcelColumnAttribute)fieldNameAttribute).Heading];

                if (p.PropertyType == typeof(string))
                {
                    p.SetValue(rowInstance, GetStringValue(rowData.Value, key));
                }
                else if (p.PropertyType == typeof(DateTime?))
                {
                    p.SetValue(rowInstance, GetDateTimeValue(rowData.Value, key));
                }
                else if (p.PropertyType == typeof(int?))
                {
                    p.SetValue(rowInstance, GetIntValue(rowData.Value, key));
                }
                else
                {
                    throw new NotImplementedException("Type " + p.PropertyType + " is not supported.");
                }
            }
        }
    }

    private void WriteRow<T>(T rowInstance, ExcelRowDefinition entry, WorksheetPart worksheet, uint rowIndex)
        where T : ExcelRowDefinition
    {
        Type objtype = rowInstance.GetType();
        foreach (PropertyInfo p in objtype.GetProperties())
        {
            object? fieldNameAttribute = p.GetCustomAttributes(false).FirstOrDefault(z => z is ExcelColumnAttribute);
            if (fieldNameAttribute != null)
            {
                string columnName = ((ExcelColumnAttribute)fieldNameAttribute).Column;
                if (p.PropertyType == typeof(string))
                {
                    string? value = (string?)p.GetValue(entry);
                    worksheet.WriteValueInCell(columnName, rowIndex, value);
                }
                else if (p.PropertyType == typeof(DateTime?))
                {
                    DateTime? value = ((DateTime?)p.GetValue(entry));
                    if (value.HasValue)
                    {
                        worksheet.WriteValueInCell(columnName, rowIndex, value.Value.ToString("yyyy-MM-dd"));
                    }
                }
                else
                {
                    throw new NotImplementedException("Type " + p.PropertyType + " is not supported.");
                }
            }
        }
    }
}