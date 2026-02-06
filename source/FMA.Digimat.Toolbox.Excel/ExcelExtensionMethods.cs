using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FMA.Digimat.Toolbox.Excel.Model;
using System.Text.RegularExpressions;

namespace FMA.Digimat.Toolbox.Excel
{
    public static class ExcelExtensionMethods
    {
        public static Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetArea(this WorksheetPart self)
        {
            if (self != null)
            {
                if (self.Worksheet.SheetDimension != null)
                {
                    return self.GetRowsExcelSheetArea(self.Worksheet.SheetDimension);
                }

                var upperRightCell = GetUpperRightCell(self);
                var lowerLeftCell = GetLowerLeftCell(self);
                return self.GetRowsExcelSheetArea(upperRightCell, lowerLeftCell);
            }

            return null;
        }

        private static Cell? GetLowerLeftCell(WorksheetPart self)
        {
            var lowerLeftCell = self.Worksheet.Descendants<Row>()?.LastOrDefault()?
                .Descendants<Cell>()?.FirstOrDefault();
            if (lowerLeftCell == null)
            {
                var sheetData = self.Worksheet.ChildElements.FirstOrDefault(c => c is SheetData);
                lowerLeftCell = sheetData?.LastChild?.FirstChild as Cell;
            }
            return lowerLeftCell;
        }

        private static Cell? GetUpperRightCell(WorksheetPart self)
        {
            var upperRightCell = self.Worksheet.Descendants<Row>()?.FirstOrDefault()?
                .Descendants<Cell>()?.LastOrDefault();
            if (upperRightCell == null)
            {
                var sheetData = self.Worksheet.ChildElements.FirstOrDefault(c => c is SheetData);
                upperRightCell = sheetData?.FirstChild?.LastChild as Cell;
            }
            return upperRightCell;
        }

        private static Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetArea(this WorksheetPart self,
            Cell upperRightCell, Cell lowerLeftCell)
        {
            if (upperRightCell == null)
            {
                throw new ArgumentNullException("upperRightCell", "The provided Cell must not be null.");
            }

            if (lowerLeftCell == null)
            {
                throw new ArgumentNullException("lowerLeftCell", "The provided Cell must not be null.");
            }

            var columnStart = lowerLeftCell.GetColumnName();
            var columnEnd = upperRightCell.GetColumnName();
            var rowStart = upperRightCell.GetRowIndex();
            var rowEnd = lowerLeftCell.GetRowIndex();

            return self.GetRowsExcelSheetArea(columnStart, rowStart, columnEnd, rowEnd);
        }

        private static Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetArea(this WorksheetPart self,
            SheetDimension area)
        {
            if (area == null)
            {
                throw new ArgumentNullException("area", "The provided SheetDimension must not be null.");
            }

            return self.GetRowsExcelSheetAreaFromReference(area.Reference);
        }

        public static Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetAreaFromReference(this WorksheetPart self,
            StringValue areaReference)
        {
            if (!areaReference.HasValue || !areaReference.Value!.Contains(":"))
            {
                throw new ArgumentException("The provided areaReference must have a value and must contain a semicolon.", "areaReference");
            }

            var startEndValues = areaReference.Value.Split(':');
            var startValue = startEndValues.FirstOrDefault();
            var endValue = startEndValues.LastOrDefault();
            var columnStart = GetColumnName(startValue);
            var columnEnd = GetColumnName(endValue);
            var rowStart = GetRowIndex(startValue);
            var rowEnd = GetRowIndex(endValue);

            return self.GetRowsExcelSheetArea(columnStart, rowStart, columnEnd, rowEnd);
        }

        public static IEnumerable<Cell> GetAllCellsFromAreaReference(this WorksheetPart self, StringValue areaReference)
        {
            if (!areaReference.HasValue || !areaReference.Value!.Contains(":"))
            {
                throw new ArgumentException("The provided areaReference must have a value and must contain a semicolon.", "areaReference");
            }

            var startEndValues = areaReference.Value.Split(':');
            var startValue = startEndValues.FirstOrDefault();
            var endValue = startEndValues.LastOrDefault();
            var columnStart = GetColumnName(startValue);
            var columnEnd = GetColumnName(endValue);
            var rowStart = GetRowIndex(startValue);
            var rowEnd = GetRowIndex(endValue);

            return
                self.Worksheet.Descendants<Cell>().Where(c =>
                    c.CompareColumn(columnStart) >= 0 &&
                    c.CompareColumn(columnEnd) <= 0 &&
                    c.GetRowIndex() >= rowStart && c.GetRowIndex() <= rowEnd)
                    .OrderBy(q => q.GetRowIndex())
                    .ThenBy(r => r.GetColumnIndex());
        }

        private static Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetArea(this WorksheetPart self,
    string columnStart, uint rowStart, string columnEnd, uint rowEnd)
        {
            var returnArrayOfRows = new Dictionary<uint, Dictionary<string, string>>();
            var indexedRow = new Dictionary<string, string>();
            IEnumerable<Cell> cells =
                self.Worksheet.Descendants<Cell>().Where(
                        c =>
                            c.CellValue != null &&
                            c.CompareColumn(columnStart) >= 0 &&
                            c.CompareColumn(columnEnd) <= 0 &&
                            c.GetRowIndex() >= rowStart && c.GetRowIndex() <= rowEnd)
                    .OrderBy(q => q.GetRowIndex())
                    .ThenBy(r => r.GetColumnIndex());

            var spreadsheetDocument = self.OpenXmlPackage as SpreadsheetDocument;
            SharedStringItem[]? sharedStringItems = null;
            var shareStringPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (shareStringPart != null)
            {
                sharedStringItems = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
            }

            foreach (var cell in cells)
            {
                var columnName = cell.GetColumnName();
                var rowNumber = cell.GetRowIndex();
                var cellValue = spreadsheetDocument.GetCellValue(cell, sharedStringItems);
                Dictionary<string, string> rowInfo;
                if (!returnArrayOfRows.TryGetValue(rowNumber, out rowInfo))
                {
                    rowInfo = new Dictionary<string, string>();
                    returnArrayOfRows[rowNumber] = rowInfo;
                }

                rowInfo[columnName] = cellValue;
            }

            return returnArrayOfRows;
        }

        public static string GetCellValue(this SpreadsheetDocument self, Cell cell, SharedStringItem[]? sharedStringItems = null)
        {
            if (self == null)
            {
                throw new ArgumentNullException("self",
                    "The provided SpreadsheetDocument in the extension method must not be null.");
            }

            var foundValue = string.Empty;

            // If the content of the first cell is stored as a shared string, get the text of the first cell
            // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (sharedStringItems == null)
                {
                    var shareStringPart = self.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    sharedStringItems = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
                }
                foundValue = sharedStringItems[int.Parse(cell.CellValue.Text)].InnerText;
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

        public static int CompareColumn(this Cell cell, string comparedTo)
        {
            var c1 = cell.GetColumnIndex();
            var c2 = GetColumnIndexFromName(comparedTo);
            return c1.CompareTo(c2);
        }

        public static int CompareRow(this Cell cell, uint comparedTo)
        {
            var c1 = cell.GetRowIndex();
            return c1.CompareTo(comparedTo);
        }

        public static uint GetColumnIndex(this Cell cell)
        {
            if (cell != null && cell.CellReference.HasValue)
            {
                return GetColumnIndexFromName(GetColumnName(cell));
            }

            return 0;
        }

        public static uint GetColumnIndexFromName(this string columnName)
        {
            var result = 0.0;
            if (!string.IsNullOrWhiteSpace(columnName))
            {
                var alphabet = new[]
                {
                'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
                'U', 'V', 'W', 'X', 'Y', 'Z'
            };

                var columnNameChars = columnName.ToUpper().ToCharArray(0, columnName.Length).Reverse().ToArray();
                for (var i = 0; i < columnName.Length; i++)
                {
                    result = result + System.Math.Pow(alphabet.Length, i) *
                        (Array.IndexOf(alphabet, columnNameChars[i]) + 1);
                }
            }

            return (uint)result;
        }

        public static string GetColumnName(this Cell cell)
        {
            return GetColumnName(cell?.CellReference?.Value);
        }

        public static uint GetRowIndex(this Cell cell)
        {
            return GetRowIndex(cell?.CellReference?.Value);
        }

        public static string GetColumnName(this string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellReference);
            return match.Value;
        }

        public static string GetColumnNameFromIndex(this uint columnIndex)
        {
            var columnName = string.Empty;
            while (columnIndex > 0)
            {
                var remainder = (columnIndex - 1) % 26;
                columnName = Convert.ToChar(65 + remainder) + columnName;
                columnIndex = (columnIndex - remainder) / 26;
            }

            return columnName;
        }

        public static uint GetRowIndex(this string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            var regex = new Regex(@"\d+");
            var match = regex.Match(cellName);
            return uint.Parse(match.Value);
        }

        public static WorkbookPart GetWorkbookPart(this SpreadsheetDocument spreadsheet)
        {
            var workbookPart = spreadsheet.WorkbookPart ?? spreadsheet.AddWorkbookPart();
            if (workbookPart.Workbook is null)
            {
                workbookPart.Workbook = new Workbook(new Sheets());
                workbookPart.Workbook.Save();
            }
            return workbookPart;
        }

        public static WorksheetPart GetWorksheetPart(this WorkbookPart workbookPart)
        {
            var worksheetParts = workbookPart.GetPartsOfType<WorksheetPart>();
            var worksheetPart = worksheetParts.FirstOrDefault();
            if (worksheetPart is null)
            {
                worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                worksheetPart.Worksheet.Save();
            }
            return worksheetPart;
        }

        public static void AddSheet(this WorkbookPart workbookPart, string sheetName)
        {
            var worksheetPart = workbookPart.GetWorksheetPart();
            var newSheetId = (uint)workbookPart.Workbook.Sheets.Count() + 1;
            var sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = newSheetId,
                Name = sheetName
            };

            workbookPart.Workbook.Sheets.Append(sheet);
        }

        public static SharedStringTablePart GetSharedStringTablePart(this WorkbookPart workbookPart)
        {
            var sharedStringTableParts = workbookPart.GetPartsOfType<SharedStringTablePart>();
            var sharedStringTablePart = sharedStringTableParts.FirstOrDefault();
            if (sharedStringTablePart is null)
            {
                sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringTablePart.SharedStringTable = new SharedStringTable();
            }
            return sharedStringTablePart;
        }

        public static uint AddSharedStringItem(this WorkbookPart workbookPart, SharedStringItem value)
        {
            var sharedStringTablePart = workbookPart.GetSharedStringTablePart();
            sharedStringTablePart.SharedStringTable.AppendChild(value);
            if (sharedStringTablePart.SharedStringTable.Count is null)
            {
                sharedStringTablePart.SharedStringTable.Count = 0;
            }
            else
            {
                sharedStringTablePart.SharedStringTable.Count++;
            }
            return sharedStringTablePart.SharedStringTable.Count;
        }

        public static uint AddSharedString(this WorkbookPart workbookPart, string value)
        {
            var text = new Text(value);
            var item = new SharedStringItem(text);
            return workbookPart.AddSharedStringItem(item);
        }

        public static uint AddSharedString(this WorkbookPart workbookPart, params CellText[] values)
        {
            var elements = new List<Run>();
            foreach (var value in values)
            {
                var run = GetRun(value.Style, value.Value);
                elements.Add(run);
            }
            var item = new SharedStringItem(elements);
            return workbookPart.AddSharedStringItem(item);
        }

        public static Run GetRun(ExcelFontStyle style, string value)
        {
            var run = new Run();
            var runProperties = GetRunProperties(style);
            run.AppendChild(runProperties);
            var text = new Text(value);
            text.Space = SpaceProcessingModeValues.Preserve;
            run.AppendChild(text);
            return run;
        }

        public static RunProperties GetRunProperties(ExcelFontStyle style)
        {
            var runProperties = new RunProperties();
            if (string.IsNullOrEmpty(style.FontColor) == false)
            {
                runProperties.AppendChild(new Color { Rgb = style.FontColor });
            }
            if (style.IsExtend.HasValue)
            {
                runProperties.AppendChild(new Extend { Val = style.IsExtend.Value });
            }
            if (style.IsCondense.HasValue)
            {
                runProperties.AppendChild(new Condense { Val = style.IsCondense.Value });
            }
            if (style.IsItalic.HasValue)
            {
                runProperties.AppendChild(new Italic { Val = style.IsItalic.Value });
            }
            if (style.IsBold.HasValue)
            {
                runProperties.AppendChild(new Bold { Val = style.IsBold.Value });
            }
            if (string.IsNullOrEmpty(style.FontName) == false)
            {
                runProperties.AppendChild(new RunFont { Val = style.FontName });
            }
            if (style.FontSize.HasValue)
            {
                runProperties.AppendChild(new FontSize { Val = style.FontSize.Value });
            }
            if (style.HasShadow.HasValue)
            {
                runProperties.AppendChild(new Shadow { Val = style.HasShadow.Value });
            }
            if (style.IsStrike.HasValue)
            {
                runProperties.AppendChild(new Strike { Val = style.IsStrike.Value });
            }
            if (style.UnderlineType.HasValue)
            {
                runProperties.AppendChild(new Underline { Val = style.UnderlineType.Value });
            }
            if (style.VerticalAlignment.HasValue)
            {
                runProperties.AppendChild(new VerticalTextAlignment { Val = style.VerticalAlignment.Value });
            }
            return runProperties;
        }
    }
}