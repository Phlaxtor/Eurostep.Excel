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

                Cell? upperRightCell = GetUpperRightCell(self);
                Cell? lowerLeftCell = GetLowerLeftCell(self);
                return self.GetRowsExcelSheetArea(upperRightCell, lowerLeftCell);
            }

            return null;
        }

        private static Cell? GetLowerLeftCell(WorksheetPart self)
        {
            Cell? lowerLeftCell = self.Worksheet.Descendants<Row>()?.LastOrDefault()?
                .Descendants<Cell>()?.FirstOrDefault();
            if (lowerLeftCell == null)
            {
                OpenXmlElement? sheetData = self.Worksheet.ChildElements.FirstOrDefault(c => c is SheetData);
                lowerLeftCell = sheetData?.LastChild?.FirstChild as Cell;
            }
            return lowerLeftCell;
        }

        private static Cell? GetUpperRightCell(WorksheetPart self)
        {
            Cell? upperRightCell = self.Worksheet.Descendants<Row>()?.FirstOrDefault()?
                .Descendants<Cell>()?.LastOrDefault();
            if (upperRightCell == null)
            {
                OpenXmlElement? sheetData = self.Worksheet.ChildElements.FirstOrDefault(c => c is SheetData);
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

            string columnStart = lowerLeftCell.GetColumnName();
            string columnEnd = upperRightCell.GetColumnName();
            uint rowStart = upperRightCell.GetRowIndex();
            uint rowEnd = lowerLeftCell.GetRowIndex();

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

            string[] startEndValues = areaReference.Value.Split(':');
            string? startValue = startEndValues.FirstOrDefault();
            string? endValue = startEndValues.LastOrDefault();
            string columnStart = GetColumnName(startValue);
            string columnEnd = GetColumnName(endValue);
            uint rowStart = GetRowIndex(startValue);
            uint rowEnd = GetRowIndex(endValue);

            return self.GetRowsExcelSheetArea(columnStart, rowStart, columnEnd, rowEnd);
        }

        public static IEnumerable<Cell> GetAllCellsFromAreaReference(this WorksheetPart self, StringValue areaReference)
        {
            if (!areaReference.HasValue || !areaReference.Value!.Contains(":"))
            {
                throw new ArgumentException("The provided areaReference must have a value and must contain a semicolon.", "areaReference");
            }

            string[] startEndValues = areaReference.Value.Split(':');
            string? startValue = startEndValues.FirstOrDefault();
            string? endValue = startEndValues.LastOrDefault();
            string columnStart = GetColumnName(startValue);
            string columnEnd = GetColumnName(endValue);
            uint rowStart = GetRowIndex(startValue);
            uint rowEnd = GetRowIndex(endValue);

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

            SpreadsheetDocument? spreadsheetDocument = self.OpenXmlPackage as SpreadsheetDocument;
            SharedStringItem[]? sharedStringItems = null;
            SharedStringTablePart? shareStringPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (shareStringPart != null)
            {
                sharedStringItems = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
            }

            foreach (Cell cell in cells)
            {
                string columnName = cell.GetColumnName();
                uint rowNumber = cell.GetRowIndex();
                string cellValue = spreadsheetDocument.GetCellValue(cell, sharedStringItems);
                if (!returnArrayOfRows.TryGetValue(rowNumber, out Dictionary<string, string> rowInfo))
                {
                    rowInfo = [];
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

            string foundValue = string.Empty;

            // If the content of the first cell is stored as a shared string, get the text of the first cell
            // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (sharedStringItems == null)
                {
                    SharedStringTablePart shareStringPart = self.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
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
            uint c1 = cell.GetColumnIndex();
            uint c2 = GetColumnIndexFromName(comparedTo);
            return c1.CompareTo(c2);
        }

        public static int CompareRow(this Cell cell, uint comparedTo)
        {
            uint c1 = cell.GetRowIndex();
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
            double result = 0.0;
            if (!string.IsNullOrWhiteSpace(columnName))
            {
                char[] alphabet = new[]
                {
                'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
                'U', 'V', 'W', 'X', 'Y', 'Z'
            };

                char[] columnNameChars = columnName.ToUpper().ToCharArray(0, columnName.Length).Reverse().ToArray();
                for (int i = 0; i < columnName.Length; i++)
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
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        }

        public static string GetColumnNameFromIndex(this uint columnIndex)
        {
            string columnName = string.Empty;
            while (columnIndex > 0)
            {
                uint remainder = (columnIndex - 1) % 26;
                columnName = Convert.ToChar(65 + remainder) + columnName;
                columnIndex = (columnIndex - remainder) / 26;
            }

            return columnName;
        }

        public static uint GetRowIndex(this string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);
            return uint.Parse(match.Value);
        }

        public static WorkbookPart GetWorkbookPart(this SpreadsheetDocument spreadsheet)
        {
            WorkbookPart workbookPart = spreadsheet.WorkbookPart ?? spreadsheet.AddWorkbookPart();
            if (workbookPart.Workbook is null)
            {
                workbookPart.Workbook = new Workbook(new Sheets());
                workbookPart.Workbook.Save();
            }
            return workbookPart;
        }

        public static WorksheetPart GetWorksheetPart(this WorkbookPart workbookPart)
        {
            IEnumerable<WorksheetPart> worksheetParts = workbookPart.GetPartsOfType<WorksheetPart>();
            WorksheetPart? worksheetPart = worksheetParts.FirstOrDefault();
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
            WorksheetPart worksheetPart = workbookPart.GetWorksheetPart();
            uint newSheetId = (uint)workbookPart.Workbook.Sheets.Count() + 1;
            Sheet sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = newSheetId,
                Name = sheetName
            };

            workbookPart.Workbook.Sheets.Append(sheet);
        }

        public static SharedStringTablePart GetSharedStringTablePart(this WorkbookPart workbookPart)
        {
            IEnumerable<SharedStringTablePart> sharedStringTableParts = workbookPart.GetPartsOfType<SharedStringTablePart>();
            SharedStringTablePart? sharedStringTablePart = sharedStringTableParts.FirstOrDefault();
            if (sharedStringTablePart is null)
            {
                sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringTablePart.SharedStringTable = new SharedStringTable();
            }
            return sharedStringTablePart;
        }

        public static uint AddSharedStringItem(this WorkbookPart workbookPart, SharedStringItem value)
        {
            SharedStringTablePart sharedStringTablePart = workbookPart.GetSharedStringTablePart();
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
            Text text = new Text(value);
            SharedStringItem item = new SharedStringItem(text);
            return workbookPart.AddSharedStringItem(item);
        }

        public static uint AddSharedString(this WorkbookPart workbookPart, params CellText[] values)
        {
            List<Run> elements = [];
            foreach (CellText value in values)
            {
                Run run = GetRun(value.Style, value.Value);
                elements.Add(run);
            }
            SharedStringItem item = new SharedStringItem(elements);
            return workbookPart.AddSharedStringItem(item);
        }

        public static Run GetRun(ExcelFontStyle style, string value)
        {
            Run run = new Run();
            RunProperties runProperties = GetRunProperties(style);
            run.AppendChild(runProperties);
            Text text = new Text(value)
            {
                Space = SpaceProcessingModeValues.Preserve
            };
            run.AppendChild(text);
            return run;
        }

        public static RunProperties GetRunProperties(ExcelFontStyle style)
        {
            RunProperties runProperties = new RunProperties();
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