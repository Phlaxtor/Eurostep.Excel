using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FMA.Digimat.Toolbox.Excel.Model;

namespace FMA.Digimat.Toolbox.Excel
{
    public static partial class ExcelGenerationHelperMethods
    {
        public static Table GetTable(this WorksheetPart self, string tableName)
        {
            var table = self.TableDefinitionParts.FirstOrDefault(r => r.Table.Name == tableName);
            if (table == null)
            {
                throw new ApplicationException($"Table {tableName} not found.");
            }

            return table.Table;
        }

        public static Dictionary<uint, Dictionary<string, string>> GetTableRows(this WorksheetPart self, string tableName)
        {
            var table = self.GetTable(tableName);

            return self.GetTableRows(table);
        }

        public static Dictionary<uint, Dictionary<string, string>> GetTableRows(this WorksheetPart self, Table table)
        {
            return self.GetRowsExcelSheetAreaFromReference(table.Reference);
        }

        public static WorksheetPart GetWorksheetPartBySheetId(this SpreadsheetDocument self, string sheetId)
        {
            return (WorksheetPart)self.WorkbookPart.GetPartById(sheetId);
        }

        public static WorksheetPart GetWorksheetPartBySheetName(this SpreadsheetDocument self, string sheetName)
        {
            var sheet = self.WorkbookPart.Workbook.Sheets.FirstOrDefault(s => ((Sheet)s).Name == sheetName) as Sheet;
            if (sheet == null)
            {
                throw new ApplicationException($"Sheet {sheetName} not found");
            }
            return (WorksheetPart)self.WorkbookPart.GetPartById(sheet.Id);
        }

        public static Column AddColumn(this WorksheetPart worksheetPart, uint columnIndex, double width)
        {
            var column = CreateColumn(columnIndex, width);
            var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
            if (columns == null)
            {
                columns = new Columns();
                InsertChildElement(worksheetPart.Worksheet, columns);
            }

            columns.Append(column);
            return column;
        }

        public static Cell AddRowWithCell(this WorkbookPart workbookPart, string columnName, uint rowIndex, params CellText[] values)
        {
            var worksheetPart = workbookPart.GetWorksheetPart();
            var row = worksheetPart.AddRow(rowIndex);
            var index = workbookPart.AddSharedString(values);
            var cell = row.AddCell(columnName, dataType: CellValues.SharedString);
            cell.CellValue = new CellValue(index.ToString());
            return cell;
        }

        public static Cell AddRowWithCell(this WorkbookPart workbookPart, string columnName, uint rowIndex, string value, double? height = null, uint? styleIndex = null)
        {
            var worksheetPart = workbookPart.GetWorksheetPart();
            var row = worksheetPart.AddRow(rowIndex, height);
            return row.AddCell(columnName, value, styleIndex, CellValues.String);
        }

        public static Cell AddCell(this Row row, string columnName, string? value = null, uint? styleIndex = null, CellValues dataType = CellValues.String)
        {
            var cell = new Cell();
            cell.CellReference = string.Format("{0}{1}", columnName, row.RowIndex.Value);
            cell.DataType = new EnumValue<CellValues>(dataType);
            if (styleIndex.HasValue)
            {
                cell.StyleIndex = styleIndex.Value;
            }
            if (string.IsNullOrEmpty(value) == false)
            {
                cell.CellValue = new CellValue(value);
            }
            row.AppendChild(cell);
            return cell;
        }

        public static Row AddRow(this WorksheetPart worksheetPart, uint rowIndex, double? height = null)
        {
            var row = new Row() { RowIndex = rowIndex };
            if (height.HasValue)
            {
                row.Height = height.Value;
                row.CustomHeight = true;
            }
            worksheetPart.Worksheet.GetFirstChild<SheetData>().AppendChild(row);
            return row;
        }

        public static void SetSheetDimension(this WorksheetPart worksheetPart, string lastColumn, uint lastRow, string firstColumn = "A", uint firstRow = 1)
        {
            var sheetDimension = new SheetDimension();
            sheetDimension.Reference = new StringValue($"{firstColumn}{firstRow}:{lastColumn}{lastRow}");
            worksheetPart.Worksheet.SheetDimension = sheetDimension;
        }

        public static void AddFilter(this WorksheetPart worksheetPart, string firstColumn, uint firstRow, string lastColumn, uint lastRow)
        {
            var autoFilter = new AutoFilter();
            autoFilter.Reference = new StringValue($"{firstColumn}{firstRow}:{lastColumn}{lastRow}");
            worksheetPart.Worksheet.InsertChildElement(autoFilter);
        }

        public static void MergeRowCells(this WorksheetPart worksheetPart, uint row, uint firstCol, uint lastCol)
        {
            var reference = $"{firstCol.GetColumnNameFromIndex()}{row}:{lastCol.GetColumnNameFromIndex()}{row}";
            worksheetPart.MergeCells(reference);
        }

        public static void MergeCells(this WorksheetPart worksheetPart, string mergeCellsReference)
        {
            var mergeCells = worksheetPart.Worksheet.GetFirstChild<MergeCells>();
            if (mergeCells == null)
            {
                mergeCells = new MergeCells();
                InsertChildElement(worksheetPart.Worksheet, mergeCells);
            }
            var mergeCell = new MergeCell();
            mergeCell.Reference = mergeCellsReference;
            mergeCells.Append(mergeCell);
        }

        public static void WriteStringValueInCell(this WorksheetPart self, uint columnNo, uint rowIndex, string cellValue,
            uint? styleIndex = null)
        {
            var columnName = columnNo.GetColumnNameFromIndex();
            var cell = self.GetCell(columnName, rowIndex);
            var value = string.IsNullOrWhiteSpace(cellValue) ? string.Empty : cellValue;
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
            cell.CellValue = new CellValue(value);
            if (styleIndex.HasValue)
            {
                cell.StyleIndex = UInt32Value.ToUInt32(styleIndex.Value);
            }
        }

        public static Cell GetCell(this WorksheetPart self, string columnName, uint rowIndex,
            CellValues dataType = CellValues.String, uint? styleIndex = null)
        {
            if (string.IsNullOrWhiteSpace(columnName))
            {
                throw new ArgumentNullException("columnName",
                    "The provided value for the column must not be null empty or contain whitespaces only.");
            }

            if (self == null)
            {
                throw new ArgumentNullException("self", "The provided WorksheetPart must not be null.");
            }

            var sheetData = self.Worksheet.GetFirstChild<SheetData>();
            var cellReference = string.Format("{0}{1}", columnName, rowIndex);

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }

            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            var columnIndex = columnName.GetColumnIndexFromName();
            foreach (var cell in row.Elements<Cell>())
            {
                var cellIndex = cell.GetColumnIndex();
                if (cellIndex > columnIndex)
                {
                    refCell = cell;
                    break;
                }
            }

            var newCell = new Cell();
            newCell.CellReference = StringValue.ToString(cellReference);
            newCell.DataType = new EnumValue<CellValues>(dataType);
            if (styleIndex.HasValue)
            {
                newCell.StyleIndex = styleIndex.Value;
            }

            row.InsertBefore(newCell, refCell);
            return newCell;
        }

        private static Column CreateColumn(uint columnIndex, double columnWidth)
        {
            var column = new Column();
            column.Min = new UInt32Value(columnIndex);
            column.Max = new UInt32Value(columnIndex);
            column.Width = new DoubleValue(columnWidth);
            column.CustomWidth = new BooleanValue(true);
            column.BestFit = new BooleanValue(true);
            return column;
        }

        private static void AddDataValidation(this Worksheet worksheet, DataValidation dataValidation)
        {
            var dataValidations = worksheet.GetFirstChild<DataValidations>();
            if (dataValidations == null)
            {
                dataValidations = new DataValidations();
                worksheet.InsertChildElement(dataValidations);
            }

            dataValidations.Append(dataValidation);
        }

        public static void AddListValidation(this WorksheetPart worksheetPart, string validatedCellsRange, string formula1Text, string errorTitle, string errorMessage = null)
        {
            worksheetPart.AddDataValidation(DataValidationValues.List, validatedCellsRange, formula1Text, null, true, errorTitle, errorMessage);
        }

        public static void AddCustomValidation(this WorksheetPart worksheetPart, string validatedCellsRange, string formula1Text, string errorTitle, string errorMessage)
        {
            worksheetPart.AddDataValidation(DataValidationValues.Custom, validatedCellsRange, formula1Text, null, true, errorTitle, errorMessage);
        }

        public static void AddTextLengthValidation(this WorksheetPart worksheetPart, string validatedCellsRange, string minLength, string maxLength, string errorTitle, string errorMessage)
        {
            worksheetPart.AddDataValidation(DataValidationValues.TextLength, validatedCellsRange, minLength, maxLength, true, errorTitle, errorMessage);
        }

        public static void AddDataValidation(this WorksheetPart worksheetPart, DataValidationValues validationType, string validatedCellsRange, string formula1Text, string formula2Text, bool allowBlanks, string errorTitle, string errorText)
        {
            var validation = new DataValidation
            {
                AllowBlank = allowBlanks,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = validatedCellsRange },
                Type = validationType
            };

            if (errorText != null || errorTitle != null)
            {
                validation.ShowErrorMessage = true;
                validation.ErrorTitle = errorTitle;
                validation.Error = errorText;
            }
            if (formula1Text != null)
            {
                var formula = new Formula1 { Text = formula1Text };
                validation.Append(formula);
            }
            if (formula2Text != null)
            {
                var formula = new Formula2 { Text = formula2Text };
                validation.Append(formula);
            }
            worksheetPart.Worksheet.AddDataValidation(validation);
        }

        public static void AddInputPrompt(this WorksheetPart worksheetPart, string validatedCellsRange, string promptTitle, string promptText)
        {
            if (promptTitle.Length > 32) throw new ArgumentException($"Title too long (max 32 characters allowed, actual length: {promptTitle.Length})", "promptTitle");
            if (promptText.Length > 255) throw new ArgumentException($"Prompt text too long (max 255 characters allowed, actual length: {promptText.Length})", "promptText");
            promptText = promptText.Replace("\\r\\n", "_x000a_");
            promptText = promptText.Replace("\r\n", "_x000a_");
            var validation = new DataValidation
            {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = validatedCellsRange },
                ShowInputMessage = true,
                PromptTitle = promptTitle,
                Prompt = promptText
            };
            worksheetPart.Worksheet.AddDataValidation(validation);
        }

        public static WorkbookView GetWorkbookView(this Workbook self)
        {
            if (self.BookViews == null)
            {
                self.BookViews = new BookViews();
            }
            var view = self.BookViews.ChildElements.First<WorkbookView>();
            if (view == null)
            {
                view = self.BookViews.AppendChild(new WorkbookView());
            }
            return view;
        }

        public static Metadata AddMetadata(this WorkbookPart self)
        {
            if (self.CellMetadataPart == null)
            {
                self.AddNewPart<CellMetadataPart>();
                self.CellMetadataPart.Metadata = new("<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xlrd=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\">\r\n\t<metadataTypes count=\"2\">\r\n\t\t<metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/>\r\n\t\t<metadataType name=\"XLRICHVALUE\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\"/>\r\n\t</metadataTypes>\r\n\t<futureMetadata name=\"XLDAPR\" count=\"1\">\r\n\t\t<bk>\r\n\t\t\t<extLst>\r\n\t\t\t\t<ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\">\r\n\t\t\t\t\t<xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/>\r\n\t\t\t\t</ext>\r\n\t\t\t</extLst>\r\n\t\t</bk>\r\n\t</futureMetadata>\r\n\t<futureMetadata name=\"XLRICHVALUE\" count=\"2\">\r\n\t\t<bk>\r\n\t\t\t<extLst>\r\n\t\t\t\t<ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\">\r\n\t\t\t\t\t<xlrd:rvb i=\"0\"/>\r\n\t\t\t\t</ext>\r\n\t\t\t</extLst>\r\n\t\t</bk>\r\n\t\t<bk>\r\n\t\t\t<extLst>\r\n\t\t\t\t<ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\">\r\n\t\t\t\t\t<xlrd:rvb i=\"1\"/>\r\n\t\t\t\t</ext>\r\n\t\t\t</extLst>\r\n\t\t</bk>\r\n\t</futureMetadata>\r\n\t<cellMetadata count=\"1\">\r\n\t\t<bk>\r\n\t\t\t<rc t=\"1\" v=\"0\"/>\r\n\t\t</bk>\r\n\t</cellMetadata>\r\n\t<valueMetadata count=\"2\">\r\n\t\t<bk>\r\n\t\t\t<rc t=\"2\" v=\"0\"/>\r\n\t\t</bk>\r\n\t\t<bk>\r\n\t\t\t<rc t=\"2\" v=\"1\"/>\r\n\t\t</bk>\r\n\t</valueMetadata>\r\n</metadata>");
            }
            return self.CellMetadataPart.Metadata;
        }

        public static void InsertChildElement(this Worksheet sheet, OpenXmlElement child)
        {
            var sheetDataPosition = 5;
            // NB: Worksheet children must be appended in the correct order (matching the order in the sequence of CT_Worksheet in sml.xsd)
            // we can assumme that SheetData is always present
            Type[] sequence =
            {
            typeof(SheetProperties), //<xsd:element name="sheetPr" type="CT_SheetPr" minOccurs="0" maxOccurs="1"/>
            typeof(Dimension), //<xsd:element name="dimension" type="CT_SheetDimension" minOccurs="0" maxOccurs="1"/>
            typeof(SheetViews), //<xsd:element name="sheetViews" type="CT_SheetViews" minOccurs="0" maxOccurs="1"/>
            typeof(SheetFormatProperties), //<xsd:element name="sheetFormatPr" type="CT_SheetFormatPr" minOccurs="0" maxOccurs="1"/>
            typeof(Columns), //<xsd:element name="cols" type="CT_Cols" minOccurs="0" maxOccurs="unbounded"/>
            typeof(SheetData), //<xsd:element name="sheetData" type="CT_SheetData" minOccurs="1" maxOccurs="1"/>
            typeof(SheetCalculationProperties), //<xsd:element name="sheetCalcPr" type="CT_SheetCalcPr" minOccurs="0" maxOccurs="1"/>
            typeof(SheetProtection), //<xsd:element name="sheetProtection" type="CT_SheetProtection" minOccurs="0" maxOccurs="1"/>
            typeof(ProtectedRanges), //<xsd:element name="protectedRanges" type="CT_ProtectedRanges" minOccurs="0" maxOccurs="1"/>
            typeof(Scenarios), //<xsd:element name="scenarios" type="CT_Scenarios" minOccurs="0" maxOccurs="1"/>
            typeof(AutoFilter), //<xsd:element name="autoFilter" type="CT_AutoFilter" minOccurs="0" maxOccurs="1"/>
            typeof(SortState), //<xsd:element name="sortState" type="CT_SortState" minOccurs="0" maxOccurs="1"/>
            typeof(DataConsolidate), //<xsd:element name="dataConsolidate" type="CT_DataConsolidate" minOccurs="0" maxOccurs="1"/>
            typeof(CustomSheetViews), //<xsd:element name="customSheetViews" type="CT_CustomSheetViews" minOccurs="0" maxOccurs="1"/>
            typeof(MergeCells), //<xsd:element name="mergeCells" type="CT_MergeCells" minOccurs="0" maxOccurs="1"/>
            typeof(PhoneticProperties), //<xsd:element name="phoneticPr" type="CT_PhoneticPr" minOccurs="0" maxOccurs="1"/>
            typeof(ConditionalFormatting), //<xsd:element name="conditionalFormatting" type="CT_ConditionalFormatting" minOccurs="0" maxOccurs="unbounded"/>
            typeof(DataValidations), //<xsd:element name="dataValidations" type="CT_DataValidations" minOccurs="0" maxOccurs="1"/>
            typeof(Hyperlinks), //<xsd:element name="hyperlinks" type="CT_Hyperlinks" minOccurs="0" maxOccurs="1"/>
            typeof(PrintOptions), //<xsd:element name="printOptions" type="CT_PrintOptions" minOccurs="0" maxOccurs="1"/>
            typeof(PageMargins), //<xsd:element name="pageMargins" type="CT_PageMargins" minOccurs="0" maxOccurs="1"/>
            typeof(PageSetup), //<xsd:element name="pageSetup" type="CT_PageSetup" minOccurs="0" maxOccurs="1"/>
            typeof(HeaderFooter), //<xsd:element name="headerFooter" type="CT_HeaderFooter" minOccurs="0" maxOccurs="1"/>
            typeof(RowBreaks), //<xsd:element name="rowBreaks" type="CT_PageBreak" minOccurs="0" maxOccurs="1"/>
            typeof(ColumnBreaks), //<xsd:element name="colBreaks" type="CT_PageBreak" minOccurs="0" maxOccurs="1"/>
            typeof(CustomProperties), //<xsd:element name="customProperties" type="CT_CustomProperties" minOccurs="0" maxOccurs="1"/>
            typeof(CellWatches), //<xsd:element name="cellWatches" type="CT_CellWatches" minOccurs="0" maxOccurs="1"/>
            typeof(IgnoredErrors), //<xsd:element name="ignoredErrors" type="CT_IgnoredErrors" minOccurs="0" maxOccurs="1"/>
            // skip SmartTags since they are defined in a different library, //<xsd:element name="smartTags" type="CT_SmartTags" minOccurs="0" maxOccurs="1"/>
            typeof(Drawing), //<xsd:element name="drawing" type="CT_Drawing" minOccurs="0" maxOccurs="1"/>
            typeof(DrawingHeaderFooter), //<xsd:element name="drawingHF" type="CT_DrawingHF" minOccurs="0" maxOccurs="1"/>
            typeof(Picture), //<xsd:element name="picture" type="CT_SheetBackgroundPicture" minOccurs="0" maxOccurs="1"/>
            typeof(OleObjects), //<xsd:element name="oleObjects" type="CT_OleObjects" minOccurs="0" maxOccurs="1"/>
            typeof(Controls), //<xsd:element name="controls" type="CT_Controls" minOccurs="0" maxOccurs="1"/>
            typeof(WebPublishItems), //<xsd:element name="webPublishItems" type="CT_WebPublishItems" minOccurs="0" maxOccurs="1"/>
            typeof(TableParts), //<xsd:element name="tableParts" type="CT_TableParts" minOccurs="0" maxOccurs="1"/>
            typeof(ExtensionList) //<xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
        };

            var isBeforeSheetData = false;
            for (var i = 0; i < sheetDataPosition; i++)
            {
                if (child.GetType().Name == sequence[i].Name)
                {
                    isBeforeSheetData = true;
                    break;
                }
            }

            if (isBeforeSheetData)
            {
                InsertChildElementBefore(sheet, child, sequence);
            }
            else
            {
                InsertChildElementAfter(sheet, child, sequence);
            }
        }

        public static void FreezePanes(this WorksheetPart worksheetPart, string column, uint row)
        {
            // freeze at C5
            SheetView sheetView = new SheetView()
            {
                TabSelected = true,
                WorkbookViewId = 0
            };  //hardcoded (there is only 1)
            Pane pane = new Pane()
            {
                VerticalSplit = row - 1,
                HorizontalSplit = column.GetColumnIndexFromName() - 1,
                TopLeftCell = $"{column}{row}",
                ActivePane = PaneValues.BottomRight,
                State = PaneStateValues.Frozen
            };
            sheetView.TabSelected = true;
            sheetView.Append(pane);

            SheetViews sheetViews = new SheetViews();
            sheetViews.AppendChild(sheetView);
            worksheetPart.Worksheet.InsertChildElement(sheetViews);
        }

        public static void ProtectSheet(this WorksheetPart worksheetPart)
        {
            SheetProtection sheetProtection = new SheetProtection() { Sheet = true };
            worksheetPart.Worksheet.InsertChildElement(sheetProtection);
        }

        public static void HideSheets(this Workbook workbook, uint excludeSheetId)
        {
            WorkbookView view = workbook.GetWorkbookView();
            view.ActiveTab = excludeSheetId;
            view.FirstSheet = excludeSheetId;
            foreach (Sheet sheet in workbook.Sheets)
            {
                if (sheet.SheetId != excludeSheetId)
                {
                    sheet.State = SheetStateValues.VeryHidden;
                }
                else
                {
                    sheet.State = SheetStateValues.Visible;
                }
            }
        }

        public static Stylesheet GetStylesheet(this SpreadsheetDocument doc)
        {
            var stylesPart = doc.WorkbookPart.WorkbookStylesPart;
            if (stylesPart is null)
            {
                stylesPart = doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            }

            var stylesheet = stylesPart.Stylesheet;
            if (stylesheet is null)
            {
                stylesheet = new Stylesheet();
                stylesheet.Fonts = new Fonts();
                stylesheet.Fills = new Fills();
                stylesheet.Borders = new Borders();
                stylesheet.CellFormats = new CellFormats();
                stylesPart.Stylesheet = stylesheet;
                stylesheet.Save();
            }

            return doc.WorkbookPart.WorkbookStylesPart.Stylesheet;
        }

        private static void InsertChildElementBefore(Worksheet sheet, OpenXmlElement child, Type[] sequence)
        {
            var possiblePredecessors = GetPossibleSuccessors(child, sequence);
            new List<Type>();

            foreach (var element in sheet.ChildElements)
            {
                if (possiblePredecessors.Contains(element.GetType()))
                {
                    sheet.InsertBefore(child, element);
                    return;
                }
            }

            sheet.AppendChild(child);
        }

        private static List<Type> GetPossibleSuccessors(OpenXmlElement child, Type[] sequence)
        {
            var possibleSuccessors = new List<Type>();
            for (var i = sequence.Length - 1; i > 0; i--)
            {
                if (child.GetType().Name == sequence[i].Name)
                {
                    break;
                }

                possibleSuccessors.Add(sequence[i]);
            }

            return possibleSuccessors;
        }

        private static void InsertChildElementAfter(Worksheet sheet, OpenXmlElement child, Type[] sequence)
        {
            var possiblePredecessors = GetPossiblePredecessors(child, sequence);
            new List<Type>();

            foreach (var element in sheet.ChildElements.Reverse())
            {
                if (possiblePredecessors.Contains(element.GetType()))
                {
                    sheet.InsertAfter(child, element);
                    return;
                }
            }

            sheet.AppendChild(child);
        }

        private static List<Type> GetPossiblePredecessors(OpenXmlElement child, Type[] sequence)
        {
            var possiblePredecessors = new List<Type>();
            for (var i = 0; i < sequence.Length; i++)
            {
                if (child.GetType().Name == sequence[i].Name)
                {
                    break;
                }

                possiblePredecessors.Add(sequence[i]);
            }

            return possiblePredecessors;
        }
    }
}