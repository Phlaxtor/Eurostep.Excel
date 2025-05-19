using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPSPrio.Components.Excel;

public static class ExcelExtensionMethods
{
    public static Column AddColumn(this WorksheetPart self, uint columnNo, double columnWidth)
    {
        Column column = ExcelUtilityMethods.CreateColumn(columnNo, columnNo, columnWidth);
        Columns columns = self.Worksheet.GetFirstChild<Columns>();
        if (columns == null)
        {
            columns = new Columns();
            self.Worksheet.InsertChildElement(columns);
        }

        columns.Append(column);
        return column;
    }

    public static Column AddColumnWithHeading(this WorksheetPart self, uint columnNo, string heading, double columnWidth, uint? styleIndex = null)
    {
        Column result = self.AddColumn(columnNo, columnWidth);
        self.WriteValueInCell(columnNo, 1u, heading, CellValues.String, styleIndex);
        return result;
    }

    public static Column AddColumnWithHeadingAndDescription(this WorksheetPart self, uint columnNo, string heading, double columnWidth, string description, uint? headingStyleIndex = null, uint? descriptionStyleIndex = null)
    {
        Column result = self.AddColumnWithHeading(columnNo, heading, columnWidth, headingStyleIndex);
        self.WriteValueInCell(columnNo, 2u, description, CellValues.String, descriptionStyleIndex);
        return result;
    }

    public static void AddConditionalFormatting(this Worksheet worksheet, ConditionalFormatting conditionalFormatting)
    {
        worksheet.InsertChildElement(conditionalFormatting);
    }

    public static void AddDataValidation(this Worksheet worksheet, DataValidation dataValidation)
    {
        DataValidations dataValidations = worksheet.GetFirstChild<DataValidations>();
        if (dataValidations == null)
        {
            dataValidations = new DataValidations();
            worksheet.InsertChildElement(dataValidations);
        }

        dataValidations.Append(dataValidation);
    }

    public static void AddHyperlink(this WorksheetPart self, string columnName, uint rowIndex, string cellValue, string linkValue = null, bool isExternal = true, string hyperlinkRelationshipId = null)
    {
        if (Uri.TryCreate(linkValue ?? cellValue, UriKind.Relative, out Uri result))
        {
            Cell cell = self.GetCell(columnName, rowIndex, CellValues.InlineString);
            cell.WriteValueInCell(cellValue, CellValues.InlineString);
            HyperlinkRelationship hyperlinkRelationship = null;
            hyperlinkRelationship = ((!string.IsNullOrWhiteSpace(hyperlinkRelationshipId)) ? self.AddHyperlinkRelationship(result, isExternal, hyperlinkRelationshipId) : self.AddHyperlinkRelationship(result, isExternal));
            Hyperlinks hyperlinks = self.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
            if (hyperlinks == null)
            {
                self.Worksheet.Append(new Hyperlinks());
                hyperlinks = self.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
            }

            if (hyperlinks != null)
            {
                Hyperlink hyperlink = new Hyperlink
                {
                    Reference = cell.CellReference.Value,
                    Id = hyperlinkRelationship.Id
                };
                hyperlinks.Append(hyperlink);
            }
        }
    }

    public static WorksheetPart AddSheet(this SpreadsheetDocument self, string sheetName)
    {
        WorkbookPart workbookPart = self.WorkbookPart;
        Sheets sheets = self.WorkbookPart.Workbook.Sheets;
        int num = sheets.Count();
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        Sheet sheet = new Sheet
        {
            Id = self.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = (uint)(num + 1),
            Name = sheetName
        };
        sheets.Append(sheet);
        workbookPart.Workbook.Save();
        return worksheetPart;
    }

    public static void AddSheetProperties(this Worksheet worksheet, SheetProperties sheetProperties)
    {
        worksheet.InsertChildElement(sheetProperties);
    }

    public static void AddTable(this WorksheetPart self, string columnStart, uint rowStart, string columnEnd, uint rowEnd, string tableName = "Table1")
    {
        string text = $"{columnStart}{rowStart}:{columnEnd}{rowEnd}";
        string text2 = $"{columnStart}{rowStart + 1}:{columnStart}{rowEnd}";
        string text3 = $"{columnStart}{rowStart + 1}:{columnEnd}{rowEnd}";
        TableDefinitionPart tableDefinitionPart = self.GetPartsOfType<TableDefinitionPart>().FirstOrDefault();
        if (tableDefinitionPart == null)
        {
            tableDefinitionPart = self.AddNewPart<TableDefinitionPart>("rId1");
        }

        Table table = new Table
        {
            Id = 1u,
            Name = tableName,
            DisplayName = tableName,
            Reference = text,
            TotalsRowShown = false
        };
        AutoFilter autoFilter = new AutoFilter
        {
            Reference = StringValue.ToString(text)
        };
        SortCondition sortCondition = new SortCondition
        {
            Reference = StringValue.ToString(text2),
            Descending = BooleanValue.ToBoolean(true)
        };
        SortState sortState = new SortState
        {
            Reference = StringValue.ToString(text3)
        };
        sortState.Append(sortCondition);
        uint columnIndexFromName = ExcelUtilityMethods.GetColumnIndexFromName(columnStart);
        uint columnIndexFromName2 = ExcelUtilityMethods.GetColumnIndexFromName(columnEnd);
        TableColumns tableColumns = new TableColumns
        {
            Count = columnIndexFromName2 - columnIndexFromName
        };
        for (uint num = columnIndexFromName; num <= columnIndexFromName2; num++)
        {
            string columnNameFromIndex = ExcelUtilityMethods.GetColumnNameFromIndex(num);
            string cellValue = self.GetCellValue(columnNameFromIndex, rowStart);
            TableColumn tableColumn = new TableColumn
            {
                Id = num,
                Name = cellValue
            };
            tableColumns.Append(tableColumn);
        }

        TableStyleInfo tableStyleInfo = new TableStyleInfo
        {
            Name = "TableStyleMedium2",
            ShowFirstColumn = false,
            ShowLastColumn = false,
            ShowRowStripes = true,
            ShowColumnStripes = false
        };
        table.Append(autoFilter);
        table.Append(sortState);
        table.Append(tableColumns);
        table.Append(tableStyleInfo);
        tableDefinitionPart.Table = table;
        TableParts tableParts = new TableParts
        {
            Count = 1u
        };
        TablePart tablePart = new TablePart
        {
            Id = "rId1"
        };
        tableParts.Append(tablePart);
        self.Worksheet.Append(tableParts);
        DifferentialFormats differentialFormats = (self.OpenXmlPackage as SpreadsheetDocument).WorkbookPart.WorkbookStylesPart.RootElement.Descendants<DifferentialFormats>().FirstOrDefault();
        if (differentialFormats != null)
        {
            DifferentialFormats differentialFormats2 = new DifferentialFormats
            {
                Count = 1u
            };
            DifferentialFormat differentialFormat = new DifferentialFormat();
            NumberingFormat numberingFormat = new NumberingFormat
            {
                NumberFormatId = 10u,
                FormatCode = "\"$\"#,##0_);[Red]\\(\"$\"#,##0\\)"
            };
            differentialFormat.Append(numberingFormat);
            differentialFormats2.Append(differentialFormat);
            new TableStyles
            {
                Count = 0u,
                DefaultTableStyle = "TableStyleMedium2",
                DefaultPivotStyle = "PivotStyleMedium9"
            };
            differentialFormats.Parent.ReplaceChild(differentialFormats2, differentialFormats);
        }
    }

    public static int CompareColumn(this Cell cell, string comparedTo)
    {
        uint columnIndex = cell.GetColumnIndex();
        uint columnIndexFromName = ExcelUtilityMethods.GetColumnIndexFromName(comparedTo);
        return columnIndex.CompareTo(columnIndexFromName);
    }

    public static int CompareRow(this Cell cell, uint comparedTo)
    {
        return cell.GetRowIndex().CompareTo(comparedTo);
    }

    public static uint? CreateBorder(this Stylesheet self, BorderStyleValues? style, System.Drawing.Color color)
    {
        return self.CreateBorder(style, color, style, color, style, color, style, color);
    }

    public static uint? CreateBorder(this Stylesheet self, BorderStyleValues style, HexBinaryValue argbColor)
    {
        Border border = new Border
        {
            TopBorder = new TopBorder
            {
                Style = new EnumValue<BorderStyleValues>(style),
                Color = new DocumentFormat.OpenXml.Spreadsheet.Color
                {
                    Rgb = argbColor
                }
            },
            RightBorder = new RightBorder
            {
                Style = new EnumValue<BorderStyleValues>(style),
                Color = new DocumentFormat.OpenXml.Spreadsheet.Color
                {
                    Rgb = argbColor
                }
            },
            BottomBorder = new BottomBorder
            {
                Style = new EnumValue<BorderStyleValues>(style),
                Color = new DocumentFormat.OpenXml.Spreadsheet.Color
                {
                    Rgb = argbColor
                }
            },
            LeftBorder = new LeftBorder
            {
                Style = new EnumValue<BorderStyleValues>(style),
                Color = new DocumentFormat.OpenXml.Spreadsheet.Color
                {
                    Rgb = argbColor
                }
            }
        };
        self.Borders.Append(border);
        Borders? borders = self.Borders;
        UInt32Value count = borders.Count;
        borders.Count = (uint)count + 1;
        UInt32Value? count2 = self.Borders.Count;
        self.Save();
        return count2;
    }

    public static uint? CreateBorder(this Stylesheet self, BorderStyleValues? top, System.Drawing.Color topColor, BorderStyleValues? right, System.Drawing.Color rightColor, BorderStyleValues? bottom, System.Drawing.Color bottomColor, BorderStyleValues? left, System.Drawing.Color leftColor)
    {
        if (self == null)
        {
            throw new ArgumentNullException("self", "The provided Stylesheet in the extension method must not be null.");
        }

        if (self.Borders == null)
        {
            throw new ApplicationException("Stylesheet.Borders must not be null.");
        }

        Border border = new Border();
        if (top.HasValue)
        {
            border.TopBorder = new TopBorder
            {
                Style = new EnumValue<BorderStyleValues>(top.Value),
                Color = ExcelUtilityMethods.GetSpreadsheetColor<DocumentFormat.OpenXml.Spreadsheet.Color>(topColor)
            };
        }

        if (right.HasValue)
        {
            border.RightBorder = new RightBorder
            {
                Style = new EnumValue<BorderStyleValues>(right.Value),
                Color = ExcelUtilityMethods.GetSpreadsheetColor<DocumentFormat.OpenXml.Spreadsheet.Color>(rightColor)
            };
        }

        if (bottom.HasValue)
        {
            border.BottomBorder = new BottomBorder
            {
                Style = new EnumValue<BorderStyleValues>(bottom.Value),
                Color = ExcelUtilityMethods.GetSpreadsheetColor<DocumentFormat.OpenXml.Spreadsheet.Color>(bottomColor)
            };
        }

        if (left.HasValue)
        {
            border.LeftBorder = new LeftBorder
            {
                Style = new EnumValue<BorderStyleValues>(left.Value),
                Color = ExcelUtilityMethods.GetSpreadsheetColor<DocumentFormat.OpenXml.Spreadsheet.Color>(leftColor)
            };
        }

        self.Borders.Append(border);
        Borders? borders = self.Borders;
        UInt32Value count = borders.Count;
        borders.Count = (uint)count + 1;
        UInt32Value? count2 = self.Borders.Count;
        self.Save();
        return count2;
    }

    public static uint? CreateCellFormat(this Stylesheet self, uint? numberFormatId, uint? formatId, Alignment alignment, uint? fontIndex, uint? borderId, uint? fillIndex, Protection protection, bool? pivotButton, bool? quotePrefix)
    {
        if (self == null)
        {
            throw new ArgumentNullException("self", "The provided Stylesheet in the extension method must not be null.");
        }

        if (self.CellFormats == null)
        {
            throw new ApplicationException("Stylesheet.CellFormats must not be null.");
        }

        CellFormat cellFormat = new CellFormat();
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
            cellFormat.ApplyProtection = BooleanValue.FromBoolean(value: true);
        }

        if (formatId.HasValue)
        {
            cellFormat.FormatId = UInt32Value.FromUInt32(formatId.Value);
        }

        if (alignment != null)
        {
            cellFormat.Alignment = alignment;
            cellFormat.ApplyAlignment = BooleanValue.FromBoolean(value: true);
        }

        if (borderId.HasValue)
        {
            cellFormat.BorderId = UInt32Value.FromUInt32(borderId.Value);
            cellFormat.ApplyBorder = BooleanValue.FromBoolean(value: true);
        }

        if (fontIndex.HasValue)
        {
            cellFormat.FontId = UInt32Value.FromUInt32(fontIndex.Value);
            cellFormat.ApplyFont = BooleanValue.FromBoolean(value: true);
        }

        if (fillIndex.HasValue)
        {
            cellFormat.FillId = UInt32Value.FromUInt32(fillIndex.Value);
            cellFormat.ApplyFill = BooleanValue.FromBoolean(value: true);
        }

        if (numberFormatId.HasValue)
        {
            cellFormat.NumberFormatId = UInt32Value.FromUInt32(numberFormatId.Value);
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(value: true);
        }

        self.CellFormats.Append(cellFormat);
        CellFormats? cellFormats = self.CellFormats;
        UInt32Value count = cellFormats.Count;
        cellFormats.Count = (uint)count + 1;
        UInt32Value? count2 = self.CellFormats.Count;
        self.Save();
        return count2;
    }

    public static uint? CreateFill(this Stylesheet self, PatternValues patternType, System.Drawing.Color foregroundColor, System.Drawing.Color backgroundColor, GradientValues? gradientType = null, double degree = 0.0, double top = 0.0, double bottom = 0.0, double right = 0.0, double left = 0.0)
    {
        if (self == null)
        {
            throw new ArgumentNullException("self", "The provided Stylesheet in the extension method must not be null.");
        }

        if (self.Fills == null)
        {
            throw new ApplicationException("Stylesheet.Fills must not be null.");
        }

        Fill fill = new Fill
        {
            PatternFill = new PatternFill
            {
                ForegroundColor = ExcelUtilityMethods.GetSpreadsheetColor<ForegroundColor>(foregroundColor),
                BackgroundColor = ExcelUtilityMethods.GetSpreadsheetColor<BackgroundColor>(backgroundColor),
                PatternType = new EnumValue<PatternValues>(patternType)
            }
        };
        if (gradientType.HasValue)
        {
            fill.GradientFill = new GradientFill
            {
                Type = new EnumValue<GradientValues>(gradientType.Value),
                Degree = degree,
                Top = top,
                Bottom = bottom,
                Right = right,
                Left = left
            };
        }

        self.Fills.Append(fill);
        Fills? fills = self.Fills;
        UInt32Value count = fills.Count;
        fills.Count = (uint)count + 1;
        UInt32Value? count2 = self.Fills.Count;
        self.Save();
        return count2;
    }

    public static uint? CreateFont(this Stylesheet self, string fontName, double? fontSize, bool? isBold, bool? isItalic, UnderlineValues? underlineType, System.Drawing.Color color, VerticalAlignmentRunValues? verticalAlignment, bool? isStrike, bool? isCondense, bool? isExtend, bool? hasShadow)
    {
        if (self == null)
        {
            throw new ArgumentNullException("self", "The provided Stylesheet in the extension method must not be null.");
        }

        if (self.Fonts == null)
        {
            throw new ApplicationException("Stylesheet.Fonts must not be null.");
        }

        Font font = new Font();
        if (!string.IsNullOrEmpty(fontName))
        {
            font.FontName = new FontName
            {
                Val = StringValue.ToString(fontName)
            };
        }

        if (fontSize.HasValue)
        {
            font.FontSize = new FontSize
            {
                Val = DoubleValue.ToDouble(fontSize.Value)
            };
        }

        if (isBold.HasValue)
        {
            font.Bold = new Bold
            {
                Val = BooleanValue.ToBoolean(isBold.Value)
            };
        }

        if (isItalic.HasValue)
        {
            font.Italic = new Italic
            {
                Val = BooleanValue.ToBoolean(isItalic.Value)
            };
        }

        if (underlineType.HasValue)
        {
            font.Underline = new Underline
            {
                Val = new EnumValue<UnderlineValues>(underlineType.Value)
            };
        }

        if (isStrike.HasValue)
        {
            font.Strike = new Strike
            {
                Val = BooleanValue.ToBoolean(isStrike.Value)
            };
        }

        if (isCondense.HasValue)
        {
            font.Condense = new Condense
            {
                Val = BooleanValue.ToBoolean(isCondense.Value)
            };
        }

        if (isExtend.HasValue)
        {
            font.Extend = new Extend
            {
                Val = BooleanValue.ToBoolean(isExtend.Value)
            };
        }

        if (hasShadow.HasValue)
        {
            font.Shadow = new Shadow
            {
                Val = BooleanValue.ToBoolean(hasShadow.Value)
            };
        }

        if (verticalAlignment.HasValue)
        {
            font.VerticalTextAlignment = new VerticalTextAlignment
            {
                Val = new EnumValue<VerticalAlignmentRunValues>(verticalAlignment.Value)
            };
        }

        font.Color = ExcelUtilityMethods.GetSpreadsheetColor<DocumentFormat.OpenXml.Spreadsheet.Color>(color);
        self.Fonts.Append(font);
        Fonts? fonts = self.Fonts;
        UInt32Value count = fonts.Count;
        fonts.Count = (uint)count + 1;
        UInt32Value? count2 = self.Fonts.Count;
        self.Save();
        return count2;
    }

    public static uint? CreateNumberingFormat(this Stylesheet self, string formatCode, uint? numberFormatId)
    {
        if (self == null)
        {
            throw new ArgumentNullException("self", "The provided Stylesheet in the extension method must not be null.");
        }

        if (self.NumberingFormats == null)
        {
            throw new ApplicationException("Stylesheet.NumberingFormats must not be null.");
        }

        NumberingFormat numberingFormat = new NumberingFormat();
        if (!string.IsNullOrWhiteSpace(formatCode))
        {
            numberingFormat.FormatCode = StringValue.ToString(formatCode);
        }

        if (numberFormatId.HasValue)
        {
            numberingFormat.NumberFormatId = UInt32Value.ToUInt32(numberFormatId.Value);
        }

        self.NumberingFormats.Append(numberingFormat);
        NumberingFormats? numberingFormats = self.NumberingFormats;
        UInt32Value count = numberingFormats.Count;
        numberingFormats.Count = (uint)count + 1;
        UInt32Value? count2 = self.NumberingFormats.Count;
        self.Save();
        return count2;
    }

    public static uint? CreateSolidFill(this Stylesheet self, HexBinaryValue argbColor)
    {
        if (self == null)
        {
            throw new ArgumentNullException("self", "The provided Stylesheet in the extension method must not be null.");
        }

        if (self.Fills == null)
        {
            throw new ApplicationException("Stylesheet.Fills must not be null.");
        }

        Fill fill = new Fill
        {
            PatternFill = new PatternFill
            {
                PatternType = PatternValues.Solid
            }
        };
        ForegroundColor foregroundColor = new ForegroundColor
        {
            Rgb = argbColor
        };
        fill.PatternFill.Append(foregroundColor);
        fill.PatternFill.Append(new BackgroundColor
        {
            Indexed = 64u
        });
        self.Fills.Append(fill);
        Fills? fills = self.Fills;
        UInt32Value count = fills.Count;
        fills.Count = (uint)count + 1;
        UInt32Value? count2 = self.Fills.Count;
        self.Save();
        return count2;
    }

    public static Cell GetCell(this WorksheetPart self, string columnName, uint rowIndex, CellValues dataType = CellValues.String, uint? styleIndex = null)
    {
        if (string.IsNullOrWhiteSpace(columnName))
        {
            throw new ArgumentNullException("columnName", "The provided value for the column must not be null empty or contain whitespaces only.");
        }

        if (self == null)
        {
            throw new ArgumentNullException("self", "The provided WorksheetPart must not be null.");
        }

        SheetData firstChild = self.Worksheet.GetFirstChild<SheetData>();
        string cellReference = $"{columnName}{rowIndex}";
        Row row;
        if ((from r in firstChild.Elements<Row>()
             where (uint)r.RowIndex == rowIndex
             select r).Count() != 0)
        {
            row = (from r in firstChild.Elements<Row>()
                   where (uint)r.RowIndex == rowIndex
                   select r).First();
        }
        else
        {
            row = new Row
            {
                RowIndex = rowIndex
            };
            firstChild.Append(row);
        }

        if ((from c in row.Elements<Cell>()
             where c.CellReference.Value == cellReference
             select c).Count() > 0)
        {
            return (from c in row.Elements<Cell>()
                    where c.CellReference.Value == cellReference
                    select c).First();
        }

        Cell referenceChild = null;
        uint columnIndexFromName = ExcelUtilityMethods.GetColumnIndexFromName(columnName);
        foreach (Cell item in row.Elements<Cell>())
        {
            if (item.GetColumnIndex() > columnIndexFromName)
            {
                referenceChild = item;
                break;
            }
        }

        Cell cell = new Cell
        {
            CellReference = StringValue.ToString(cellReference),
            DataType = new EnumValue<CellValues>(dataType)
        };
        if (styleIndex.HasValue)
        {
            cell.StyleIndex = styleIndex.Value;
        }

        row.InsertBefore(cell, referenceChild);
        return cell;
    }

    public static string GetCellValue(this OpenXmlPackage self, Cell cell)
    {
        return self.GetSpreadsheetDocument().GetCellValue(cell);
    }

    public static string GetCellValue(this SpreadsheetDocument self, Cell cell)
    {
        if (self == null)
        {
            throw new ArgumentNullException("self", "The provided SpreadsheetDocument in the extension method must not be null.");
        }

        string result = string.Empty;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            result = self.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable.Elements<SharedStringItem>().ToArray()[int.Parse(cell.CellValue.Text)].InnerText;
        }
        else if (cell.CellValue != null)
        {
            result = cell.CellValue.Text;
        }

        return result;
    }

    public static string GetCellValue(this WorksheetPart self, string columnName, uint rowIndex)
    {
        Cell cell = self.GetCell(columnName, rowIndex);
        return self.OpenXmlPackage.GetCellValue(cell);
    }

    public static string GetCellValue(this WorksheetPart self, string cellReference)
    {
        string cellReference2 = cellReference;
        Cell cell = (from c in self.Worksheet.Descendants<Cell>()
                     where c.CellReference == cellReference2
                     select c).FirstOrDefault();
        if (cell == null)
        {
            return null;
        }

        return self.OpenXmlPackage.GetCellValue(cell);
    }

    public static Dictionary<uint, string> GetColumnData(this WorksheetPart self, string columnName)
    {
        string columnName2 = columnName;
        IOrderedEnumerable<Cell> orderedEnumerable = from c in self.Worksheet.Descendants<Cell>()
                                                     where c.GetColumnName() == columnName2
                                                     select c into r
                                                     orderby r.GetRowIndex()
                                                     select r;
        Dictionary<uint, string> dictionary = [];
        foreach (Cell item in orderedEnumerable)
        {
            dictionary[item.GetRowIndex()] = self.OpenXmlPackage.GetCellValue(item);
        }

        return dictionary;
    }

    public static uint GetColumnIndex(this Cell cell)
    {
        if (cell != null && cell.CellReference.HasValue)
        {
            return ExcelUtilityMethods.GetColumnIndexFromName(cell.GetColumnName());
        }

        return 0u;
    }

    public static string GetColumnName(this Cell cell)
    {
        if (cell != null && cell.CellReference.HasValue)
        {
            return ExcelUtilityMethods.GetColumnName(cell.CellReference.Value);
        }

        return null;
    }

    public static Dictionary<string, Dictionary<uint, string>> GetColumnsExcelSheetArea(this WorksheetPart self, bool excludeHeader = true, uint? headerRow = null)
    {
        if (self != null)
        {
            uint? headerRow2 = null;
            if (excludeHeader)
            {
                if (headerRow.HasValue)
                {
                    headerRow2 = headerRow;
                }
                else
                {
                    Cell cell = self.Worksheet.LastChild.FirstChild.FirstChild as Cell;
                    headerRow2 = cell.GetRowIndex();
                }
            }

            if (self.Worksheet.SheetDimension != null)
            {
                return self.GetColumnsExcelSheetArea(self.Worksheet.SheetDimension, excludeHeader, headerRow);
            }

            Cell upperRightCell = self.Worksheet.LastChild.FirstChild.FirstChild as Cell;
            Cell lowerLeftCell = self.Worksheet.LastChild.LastChild.LastChild as Cell;
            return self.GetColumnsExcelSheetArea(upperRightCell, lowerLeftCell, excludeHeader, headerRow2);
        }

        return null;
    }

    public static Dictionary<string, Dictionary<uint, string>> GetColumnsExcelSheetArea(this WorksheetPart self, Cell upperRightCell, Cell lowerLeftCell, bool excludeHeader = true, uint? headerRow = null)
    {
        if (upperRightCell == null)
        {
            throw new ArgumentNullException("upperRightCell", "The provided Cell must not be null.");
        }

        if (lowerLeftCell == null)
        {
            throw new ArgumentNullException("lowerLeftCell", "The provided Cell must not be null.");
        }

        string columnName = upperRightCell.GetColumnName();
        string columnName2 = lowerLeftCell.GetColumnName();
        uint num = (headerRow.HasValue ? headerRow.Value : upperRightCell.GetRowIndex());
        uint rowIndex = lowerLeftCell.GetRowIndex();
        if (excludeHeader)
        {
            num++;
        }

        return self.GetColumnsExcelSheetArea(columnName, num, columnName2, rowIndex);
    }

    public static Dictionary<string, Dictionary<uint, string>> GetColumnsExcelSheetArea(this WorksheetPart self, SheetDimension area, bool excludeHeader = true, uint? headerRow = null)
    {
        if (area == null)
        {
            throw new ArgumentNullException("area", "The provided SheetDimension must not be null.");
        }

        if (!area.Reference.HasValue)
        {
            throw new ArgumentException("The provided SheetDimension.Reference must have an value.", "area");
        }

        string[] source = area.Reference.Value.Split(':');
        string text = source.FirstOrDefault();
        string? text2 = source.LastOrDefault();
        string columnName = ExcelUtilityMethods.GetColumnName(text);
        string columnName2 = ExcelUtilityMethods.GetColumnName(text2);
        uint num = (headerRow.HasValue ? headerRow.Value : ExcelUtilityMethods.GetRowIndex(text));
        uint rowIndex = ExcelUtilityMethods.GetRowIndex(text2);
        if (excludeHeader)
        {
            num++;
        }

        return self.GetColumnsExcelSheetArea(columnName, num, columnName2, rowIndex);
    }

    public static Dictionary<string, Dictionary<uint, string>> GetColumnsExcelSheetArea(this WorksheetPart self, string columnStart, uint rowStart, string columnEnd, uint rowEnd)
    {
        string columnStart2 = columnStart;
        string columnEnd2 = columnEnd;
        Dictionary<string, Dictionary<uint, string>> dictionary = [];
        Dictionary<uint, string> dictionary2 = [];
        IOrderedEnumerable<Cell> orderedEnumerable = (from c in self.Worksheet.Descendants<Cell>()
                                                      where c.CellValue != null && c.CompareColumn(columnStart2) >= 0 && c.CompareColumn(columnEnd2) <= 0 && c.GetRowIndex() >= rowStart && c.GetRowIndex() <= rowEnd
                                                      select c into r
                                                      orderby r.GetColumnIndex()
                                                      select r).ThenBy((Cell q) => q.GetRowIndex());
        string text = string.Empty;
        foreach (Cell item in orderedEnumerable)
        {
            string columnName = item.GetColumnName();
            if (!text.Equals(columnName))
            {
                dictionary2 = (dictionary[columnName] = []);
                text = columnName;
            }

            dictionary2[item.GetRowIndex()] = self.OpenXmlPackage.GetCellValue(item);
        }

        return dictionary;
    }

    public static CellValues GetDataType(this Cell self, CellValues? dataType = null)
    {
        if (dataType.HasValue)
        {
            return dataType.Value;
        }

        return self.GetDataTypeOrDefault();
    }

    public static CellValues GetDataTypeOrDefault(this Cell self, CellValues? dataType = null)
    {
        if (self != null && self.DataType != null && self.DataType.HasValue)
        {
            return self.DataType.Value;
        }

        if (dataType.HasValue)
        {
            return dataType.Value;
        }

        return CellValues.String;
    }

    public static Cell GetFirstCellWithValue(this WorksheetPart self, string cellValue)
    {
        if (self != null)
        {
            IEnumerable<Cell> enumerable = self.Worksheet.Descendants<Cell>();
            if (enumerable != null)
            {
                foreach (Cell item in enumerable)
                {
                    string cellValue2 = self.OpenXmlPackage.GetCellValue(item);
                    if (string.Equals(cellValue.ToSingleLine(), cellValue2.ToSingleLine(), StringComparison.CurrentCulture))
                    {
                        return item;
                    }
                }
            }
        }

        return null;
    }

    public static uint GetRowIndex(this Cell cell)
    {
        if (cell != null && cell.CellReference.HasValue)
        {
            return ExcelUtilityMethods.GetRowIndex(cell.CellReference.Value);
        }

        return 0u;
    }

    public static Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetArea(this WorksheetPart self, bool excludeHeader = true, uint? headerRow = null)
    {
        if (self != null)
        {
            uint? headerRow2 = null;
            if (excludeHeader)
            {
                if (headerRow.HasValue)
                {
                    headerRow2 = headerRow;
                }
                else
                {
                    Cell cell = self.Worksheet.LastChild.FirstChild.FirstChild as Cell;
                    headerRow2 = cell.GetRowIndex();
                }
            }

            if (self.Worksheet.SheetDimension != null)
            {
                return self.GetRowsExcelSheetArea(self.Worksheet.SheetDimension, excludeHeader, headerRow);
            }

            Cell cell2 = self.Worksheet.LastChild.FirstChild.FirstChild as Cell;
            if (cell2 == null)
            {
                cell2 = self.Worksheet.Descendants<Row>()?.FirstOrDefault()?.Descendants<Cell>()?.FirstOrDefault();
            }

            Cell cell3 = self.Worksheet.LastChild.LastChild.LastChild as Cell;
            if (cell3 == null)
            {
                cell3 = self.Worksheet.Descendants<Row>()?.LastOrDefault()?.Descendants<Cell>()?.LastOrDefault();
            }

            return self.GetRowsExcelSheetArea(cell2, cell3, excludeHeader, headerRow2);
        }

        return null;
    }

    public static Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetArea(this WorksheetPart self, Cell upperRightCell, Cell lowerLeftCell, bool excludeHeader = true, uint? headerRow = null)
    {
        if (upperRightCell == null)
        {
            throw new ArgumentNullException("upperRightCell", "The provided Cell must not be null.");
        }

        if (lowerLeftCell == null)
        {
            throw new ArgumentNullException("lowerLeftCell", "The provided Cell must not be null.");
        }

        string columnName = upperRightCell.GetColumnName();
        string columnName2 = lowerLeftCell.GetColumnName();
        uint num = (headerRow.HasValue ? headerRow.Value : upperRightCell.GetRowIndex());
        uint rowIndex = lowerLeftCell.GetRowIndex();
        if (excludeHeader)
        {
            num++;
        }

        return self.GetRowsExcelSheetArea(columnName, num, columnName2, rowIndex);
    }

    public static Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetArea(this WorksheetPart self, SheetDimension area, bool excludeHeader = true, uint? headerRow = null)
    {
        if (area == null)
        {
            throw new ArgumentNullException("area", "The provided SheetDimension must not be null.");
        }

        if (!area.Reference.HasValue)
        {
            throw new ArgumentException("The provided SheetDimension.Reference must have an value.", "area");
        }

        string[] source = area.Reference.Value.Split(':');
        string text = source.FirstOrDefault();
        string? text2 = source.LastOrDefault();
        string columnName = ExcelUtilityMethods.GetColumnName(text);
        string columnName2 = ExcelUtilityMethods.GetColumnName(text2);
        uint num = (headerRow.HasValue ? headerRow.Value : ExcelUtilityMethods.GetRowIndex(text));
        uint rowIndex = ExcelUtilityMethods.GetRowIndex(text2);
        if (excludeHeader)
        {
            num++;
        }

        return self.GetRowsExcelSheetArea(columnName, num, columnName2, rowIndex);
    }

    public static Dictionary<uint, Dictionary<string, string>> GetRowsExcelSheetArea(this WorksheetPart self, string columnStart, uint rowStart, string columnEnd, uint rowEnd)
    {
        string columnStart2 = columnStart;
        string columnEnd2 = columnEnd;
        Dictionary<uint, Dictionary<string, string>> dictionary = [];
        new Dictionary<string, string>();
        foreach (Cell item in (from c in self.Worksheet.Descendants<Cell>()
                               where c.CellValue != null && c.CompareColumn(columnStart2) >= 0 && c.CompareColumn(columnEnd2) <= 0 && c.GetRowIndex() >= rowStart && c.GetRowIndex() <= rowEnd
                               select c into q
                               orderby q.GetRowIndex()
                               select q).ThenBy((Cell r) => r.GetColumnIndex()))
        {
            string columnName = item.GetColumnName();
            uint rowIndex = item.GetRowIndex();
            string cellValue = self.OpenXmlPackage.GetCellValue(item);
            if (!dictionary.TryGetValue(rowIndex, out Dictionary<string, string>? value))
            {
                value = (dictionary[rowIndex] = []);
            }

            value[columnName] = cellValue;
        }

        return dictionary;
    }

    public static SpreadsheetDocument GetSpreadsheetDocument(this OpenXmlPackage self)
    {
        return (self as SpreadsheetDocument) ?? throw new ArgumentException("Provided OpenXmlPackage is not a SpreadsheetDocument", "self");
    }

    public static Stylesheet GetStylesheet(this SpreadsheetDocument self)
    {
        return self.WorkbookPart.GetStylesheet();
    }

    public static Stylesheet GetStylesheet(this WorkbookPart self)
    {
        if (self.WorkbookStylesPart == null)
        {
            self.AddNewPart<WorkbookStylesPart>();
        }

        if (self.WorkbookStylesPart.Stylesheet == null)
        {
            Stylesheet stylesheet = new Stylesheet(new Fonts(new Font(new FontSize
            {
                Val = 11.0
            }, new DocumentFormat.OpenXml.Spreadsheet.Color
            {
                Rgb = new HexBinaryValue
                {
                    Value = "000000"
                }
            }, new FontName
            {
                Val = "Calibri"
            }))
            {
                Count = 0u
            }, new Fills(new Fill(new PatternFill
            {
                PatternType = PatternValues.None
            }), new Fill(new PatternFill
            {
                PatternType = PatternValues.Gray125
            }))
            {
                Count = 1u
            }, new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()))
            {
                Count = 0u
            }, new CellFormats(new CellFormat
            {
                FontId = 0u,
                FillId = 0u,
                BorderId = 0u
            })
            {
                Count = 0u
            });
            self.WorkbookStylesPart.Stylesheet = stylesheet;
            self.WorkbookStylesPart.Stylesheet.Save();
        }

        return self.WorkbookStylesPart.Stylesheet;
    }

    public static WorksheetPart GetWorksheetPartByCellValue(this SpreadsheetDocument self, string cellValue)
    {
        IEnumerable<Sheet> enumerable = self.WorkbookPart.Workbook.Descendants<Sheet>();
        if (enumerable != null)
        {
            foreach (Sheet item in enumerable)
            {
                WorksheetPart worksheetPartBySheetId = self.GetWorksheetPartBySheetId(item.Id);
                if (worksheetPartBySheetId == null)
                {
                    continue;
                }

                IEnumerable<Cell> enumerable2 = worksheetPartBySheetId.Worksheet.Descendants<Cell>();
                if (enumerable2 == null)
                {
                    continue;
                }

                foreach (Cell item2 in enumerable2)
                {
                    string cellValue2 = self.GetCellValue(item2);
                    if (string.Equals(cellValue, cellValue2, StringComparison.CurrentCulture))
                    {
                        return worksheetPartBySheetId;
                    }
                }
            }
        }

        return null;
    }

    public static WorksheetPart GetWorksheetPartByIndex(this SpreadsheetDocument self, int index)
    {
        IEnumerable<Sheet> enumerable = self.WorkbookPart.Workbook.Descendants<Sheet>();
        if (enumerable != null)
        {
            Sheet sheet = enumerable.ElementAtOrDefault(index);
            if (sheet != null)
            {
                return self.GetWorksheetPartBySheetId(sheet.Id);
            }
        }

        return null;
    }

    public static WorksheetPart GetWorksheetPartBySheetId(this SpreadsheetDocument self, string sheetId)
    {
        return (WorksheetPart)self.WorkbookPart.GetPartById(sheetId);
    }

    public static WorksheetPart GetWorksheetPartBySheetName(this SpreadsheetDocument self, string sheetName)
    {
        IEnumerable<Sheet> enumerable = self.WorkbookPart.Workbook.Descendants<Sheet>();
        if (enumerable != null)
        {
            foreach (Sheet item in enumerable)
            {
                if (string.Equals(sheetName, item.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    return self.GetWorksheetPartBySheetId(item.Id);
                }
            }
        }

        return null;
    }

    public static WorksheetPart InitializeSpreadsheet(this SpreadsheetDocument self, params string[] sheetNames)
    {
        WorkbookPart workbookPart = self.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        Sheets sheets = self.WorkbookPart.Workbook.AppendChild(new Sheets());
        for (int i = 0; i < sheetNames.Length; i++)
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            Sheet sheet = new Sheet
            {
                Id = self.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)(i + 1),
                Name = sheetNames[i]
            };
            sheets.Append(sheet);
        }

        workbookPart.Workbook.Save();
        workbookPart.GetStylesheet();
        return self.GetWorksheetPartBySheetName(sheetNames.First());
    }

    public static WorksheetPart InitializeSpreadsheet(this SpreadsheetDocument self, string sheetName)
    {
        WorkbookPart workbookPart = self.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        SheetData sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);
        Sheets? sheets = self.WorkbookPart.Workbook.AppendChild(new Sheets());
        Sheet sheet = new Sheet
        {
            Id = self.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1u,
            Name = sheetName
        };
        sheets.Append(sheet);
        workbookPart.Workbook.Save();
        workbookPart.GetStylesheet();
        return worksheetPart;
    }

    public static void InitializeSpreadsheetWithoutSheets(this SpreadsheetDocument self)
    {
        WorkbookPart workbookPart = self.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        self.WorkbookPart.Workbook.AppendChild(new Sheets());
        workbookPart.Workbook.Save();
        workbookPart.GetStylesheet();
    }

    public static void InsertChildElement(this Worksheet self, OpenXmlElement child)
    {
        int num = 5;
        Type[] array = new Type[36]
        {
            typeof(SheetProperties),
            typeof(Dimension),
            typeof(SheetViews),
            typeof(SheetFormatProperties),
            typeof(Columns),
            typeof(SheetData),
            typeof(SheetCalculationProperties),
            typeof(SheetProtection),
            typeof(ProtectedRanges),
            typeof(Scenarios),
            typeof(AutoFilter),
            typeof(SortState),
            typeof(DataConsolidate),
            typeof(CustomSheetViews),
            typeof(MergeCells),
            typeof(PhoneticProperties),
            typeof(ConditionalFormatting),
            typeof(DataValidations),
            typeof(Hyperlinks),
            typeof(PrintOptions),
            typeof(PageMargins),
            typeof(PageSetup),
            typeof(HeaderFooter),
            typeof(RowBreaks),
            typeof(ColumnBreaks),
            typeof(DocumentFormat.OpenXml.Spreadsheet.CustomProperties),
            typeof(CellWatches),
            typeof(IgnoredErrors),
            typeof(DocumentFormat.OpenXml.Spreadsheet.Drawing),
            typeof(DrawingHeaderFooter),
            typeof(Picture),
            typeof(OleObjects),
            typeof(Controls),
            typeof(WebPublishItems),
            typeof(TableParts),
            typeof(ExtensionList)
        };
        bool flag = false;
        for (int i = 0; i < num; i++)
        {
            if (child.GetType().Name == array[i].Name)
            {
                flag = true;
                break;
            }
        }

        if (flag)
        {
            self.InsertChildElementBefore(child, array);
        }
        else
        {
            self.InsertChildElementAfter(child, array);
        }
    }

    public static void ProtectSheet(this WorksheetPart self, bool protectSheet, bool protectObjects, bool protectScenarios)
    {
        if (self == null)
        {
            throw new ArgumentNullException("self", "The provided WorksheetPart must not be null.");
        }

        SheetData firstChild = self.Worksheet.GetFirstChild<SheetData>();
        if (firstChild == null)
        {
            self.Worksheet.AppendChild(new SheetData());
        }

        PageSetup pageSetup = self.Worksheet.GetFirstChild<PageSetup>();
        if (pageSetup == null)
        {
            SpreadsheetDocument spreadsheetDocument = self.OpenXmlPackage.GetSpreadsheetDocument();
            pageSetup = new PageSetup
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(self),
                VerticalDpi = 0u,
                Orientation = OrientationValues.Default
            };
            self.Worksheet.InsertAfter(pageSetup, firstChild);
        }

        PageMargins pageMargins = self.Worksheet.GetFirstChild<PageMargins>();
        if (pageMargins == null)
        {
            pageMargins = new PageMargins
            {
                Left = 0.7,
                Right = 0.7,
                Top = 0.75,
                Bottom = 0.75,
                Header = 0.3,
                Footer = 0.3
            };
            self.Worksheet.InsertBefore(pageMargins, pageSetup);
        }

        SheetProtection firstChild2 = self.Worksheet.GetFirstChild<SheetProtection>();
        if (firstChild2 == null)
        {
            firstChild2 = new SheetProtection
            {
                Sheet = protectSheet,
                Objects = protectObjects,
                Scenarios = protectScenarios
            };
            self.Worksheet.InsertBefore(firstChild2, pageMargins);
        }
    }

    public static void SetColumnsData(this WorksheetPart self, string columnName, double columnWidth)
    {
        Column column = ExcelUtilityMethods.CreateColumn(columnName, columnWidth);
        Columns columns = self.Worksheet.GetFirstChild<Columns>();
        if (columns == null)
        {
            columns = new Columns();
            self.Worksheet.InsertAfter(columns, self.Worksheet.SheetFormatProperties);
        }

        columns.Append(column);
    }

    public static void WriteColumnsExcelSheetArea(this WorksheetPart self, IDictionary<string, Dictionary<uint, string>> data)
    {
        if (self == null)
        {
            throw new ArgumentNullException("self", "The provided Worksheet must not be null.");
        }

        if (data == null)
        {
            return;
        }

        foreach (KeyValuePair<string, Dictionary<uint, string>> datum in data)
        {
            string key = datum.Key;
            foreach (KeyValuePair<uint, string> item in datum.Value)
            {
                uint key2 = item.Key;
                string value = item.Value;
                self.WriteValueInCell(key, key2, value);
            }
        }
    }

    public static void WriteRowsExcelSheetArea(this WorksheetPart self, IDictionary<uint, Dictionary<string, string>> data)
    {
        if (self == null)
        {
            throw new ArgumentNullException("self", "The provided Worksheet must not be null.");
        }

        if (data == null)
        {
            return;
        }

        foreach (KeyValuePair<uint, Dictionary<string, string>> datum in data)
        {
            uint key = datum.Key;
            foreach (KeyValuePair<string, string> item in datum.Value)
            {
                string key2 = item.Key;
                string value = item.Value;
                self.WriteValueInCell(key2, key, value);
            }
        }
    }

    public static Cell WriteValueInCell(this WorksheetPart self, string columnName, uint rowIndex, string? cellValue, CellValues dataType = CellValues.String, uint? styleIndex = null)
    {
        return self.GetCell(columnName, rowIndex).WriteValueInCell(cellValue, dataType, styleIndex);
    }

    public static Cell WriteValueInCell(this WorksheetPart self, uint columnNo, uint rowIndex, string cellValue, CellValues dataType = CellValues.String, uint? styleIndex = null)
    {
        string columnNameFromIndex = ExcelUtilityMethods.GetColumnNameFromIndex(columnNo);
        return self.GetCell(columnNameFromIndex, rowIndex).WriteValueInCell(cellValue, dataType, styleIndex);
    }

    public static Cell WriteValueInCell(this Cell self, string? cellValue, CellValues? dataType = null, uint? styleIndex = null)
    {
        if (self != null)
        {
            CellValues dataType2 = self.GetDataType(dataType);
            string valueForCell = ExcelUtilityMethods.GetValueForCell(cellValue, dataType2);
            dataType2 = ExcelUtilityMethods.GetDataTypeForCell(cellValue, dataType2);
            self.DataType = new EnumValue<CellValues>(dataType2);
            if (styleIndex.HasValue)
            {
                self.StyleIndex = UInt32Value.ToUInt32(styleIndex.Value);
            }

            switch (dataType2)
            {
                case CellValues.Boolean:
                    self.CellValue = new CellValue(valueForCell);
                    break;

                case CellValues.Date:
                    self.CellValue = new CellValue(valueForCell);
                    break;

                case CellValues.Error:
                    throw new NotImplementedException($"Support for CellValues value: {dataType} is not implemented yet.");
                case CellValues.InlineString:
                    {
                        Text newChild = new Text
                        {
                            Text = valueForCell
                        };
                        InlineString inlineString = new InlineString();
                        inlineString.AppendChild(newChild);
                        self.AppendChild(inlineString);
                        break;
                    }
                case CellValues.Number:
                    self.CellValue = new CellValue(valueForCell);
                    break;

                case CellValues.SharedString:
                    throw new NotImplementedException($"Support for CellValues value: {dataType} is not implemented yet.");
                case CellValues.String:
                    self.CellValue = new CellValue(valueForCell);
                    break;

                default:
                    throw new NotImplementedException($"Support for CellValues value: {dataType} is not implemented yet.");
            }
        }

        return self;
    }

    private static List<Type> GetPossiblePredecessors(OpenXmlElement child, Type[] sequence)
    {
        List<Type> possiblePredecessors = [];
        for (int i = 0; i < sequence.Length; i++)
        {
            if (child.GetType().Name == sequence[i].Name)
            {
                break;
            }
            possiblePredecessors.Add(sequence[i]);
        }
        return possiblePredecessors;
    }

    private static List<Type> GetPossibleSuccessors(OpenXmlElement child, Type[] sequence)
    {
        List<Type> possibleSuccessors = [];
        for (int i = sequence.Length - 1; i > 0; i--)
        {
            if (child.GetType().Name == sequence[i].Name)
            {
                break;
            }
            possibleSuccessors.Add(sequence[i]);
        }
        return possibleSuccessors;
    }

    private static void InsertChildElementAfter(this Worksheet self, OpenXmlElement child, Type[] sequence)
    {
        List<Type> possiblePredecessors = GetPossiblePredecessors(child, sequence); new List<Type>();

        foreach (OpenXmlElement? element in self.ChildElements.Reverse())
        {
            if (possiblePredecessors.Contains(element.GetType()))
            {
                self.InsertAfter(child, element);
                return;
            }
        }
        self.AppendChild(child);
    }

    private static void InsertChildElementBefore(this Worksheet self, OpenXmlElement child, Type[] sequence)
    {
        List<Type> possiblePredecessors = GetPossibleSuccessors(child, sequence); new List<Type>();

        foreach (OpenXmlElement element in self.ChildElements)
        {
            if (possiblePredecessors.Contains(element.GetType()))
            {
                self.InsertBefore(child, element);
                return;
            }
        }
        self.AppendChild(child);
    }
}

public static class ExcelUtilityMethods
{
    public static Column CreateColumn(uint startColumnIndex, uint endColumnIndex, double columnWidth)
    {
        return new Column
        {
            Min = new UInt32Value(startColumnIndex),
            Max = new UInt32Value(endColumnIndex),
            Width = new DoubleValue(columnWidth),
            CustomWidth = new BooleanValue(value: true),
            BestFit = new BooleanValue(value: true)
        };
    }

    public static Column CreateColumn(string columnName, double columnWidth)
    {
        uint columnIndexFromName = GetColumnIndexFromName(columnName);
        return CreateColumn(columnIndexFromName, columnIndexFromName, columnWidth);
    }

    public static Alignment GetAlignment(HorizontalAlignmentValues horizontal, uint indent, bool justifyLastLine, string mergeCell, uint readingOrder, int relativeIndent, bool shrinkToFit, uint textRotation, VerticalAlignmentValues vertical, bool wrapText)
    {
        return new Alignment
        {
            Horizontal = new EnumValue<HorizontalAlignmentValues>(horizontal),
            Indent = UInt32Value.FromUInt32(indent),
            JustifyLastLine = BooleanValue.FromBoolean(justifyLastLine),
            MergeCell = StringValue.FromString(mergeCell),
            ReadingOrder = UInt32Value.FromUInt32(readingOrder),
            RelativeIndent = Int32Value.FromInt32(relativeIndent),
            ShrinkToFit = BooleanValue.FromBoolean(shrinkToFit),
            TextRotation = UInt32Value.ToUInt32(textRotation),
            Vertical = new EnumValue<VerticalAlignmentValues>(vertical),
            WrapText = BooleanValue.FromBoolean(wrapText)
        };
    }

    public static uint GetColumnIndexFromName(string columnName)
    {
        double num = 0.0;
        if (!string.IsNullOrWhiteSpace(columnName))
        {
            char[] array = new char[26]
            {
                'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
                'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
                'U', 'V', 'W', 'X', 'Y', 'Z'
            };
            char[] array2 = columnName.ToUpper().ToCharArray(0, columnName.Length).Reverse()
                .ToArray();
            for (int i = 0; i < columnName.Length; i++)
            {
                num += System.Math.Pow(array.Length, i) * (Array.IndexOf(array, array2[i]) + 1);
            }
        }

        return (uint)num;
    }

    public static string GetColumnName(string cellReference)
    {
        return new Regex("[A-Za-z]+").Match(cellReference).Value;
    }

    public static string GetColumnNameFromIndex(uint columnIndex)
    {
        string text = string.Empty;
        while (columnIndex != 0)
        {
            uint num = (columnIndex - 1) % 26;
            text = Convert.ToChar(65 + num) + text;
            columnIndex = (columnIndex - num) / 26;
        }

        return text;
    }

    public static Protection GetProtection(bool hidden, bool locked)
    {
        return new Protection
        {
            Hidden = BooleanValue.FromBoolean(hidden),
            Locked = BooleanValue.FromBoolean(locked)
        };
    }

    public static uint GetRowIndex(string cellName)
    {
        return uint.Parse(new Regex("\\d+").Match(cellName).Value);
    }

    public static T GetSpreadsheetColor<T>(System.Drawing.Color color) where T : ColorType, new()
    {
        return new T
        {
            Rgb = new HexBinaryValue(),
            Auto = new BooleanValue(value: false)
        };
    }

    internal static CellValues GetDataTypeForCell(string cellValue, CellValues dataType = CellValues.String)
    {
        if (string.IsNullOrWhiteSpace(cellValue))
        {
            return dataType;
        }

        switch (dataType)
        {
            case CellValues.Boolean:
                if (cellValue.IsBoolean())
                {
                    return CellValues.Boolean;
                }

                return CellValues.String;

            case CellValues.Date:
                if (cellValue.IsDateTime())
                {
                    return CellValues.Date;
                }

                return CellValues.String;

            case CellValues.Error:
                return CellValues.Error;

            case CellValues.InlineString:
                return CellValues.InlineString;

            case CellValues.Number:
                if (cellValue.IsNumber())
                {
                    return CellValues.Number;
                }

                return CellValues.String;

            case CellValues.SharedString:
                return CellValues.SharedString;

            case CellValues.String:
                return CellValues.String;

            default:
                return CellValues.String;
        }
    }

    internal static string GetValueForCell(string? cellValue, CellValues dataType)
    {
        switch (dataType)
        {
            case CellValues.Boolean:
                {
                    bool? boolean = cellValue.GetBoolean();
                    if (boolean.HasValue)
                    {
                        if (!boolean.Value)
                        {
                            return "0";
                        }

                        return "1";
                    }

                    break;
                }
            case CellValues.Date:
                {
                    DateTime? dateTime = cellValue.GetDateTime();
                    if (dateTime.HasValue)
                    {
                        return Convert.ToString(dateTime.Value.ToOADate());
                    }

                    break;
                }
        }

        if (!string.IsNullOrWhiteSpace(cellValue))
        {
            return cellValue;
        }

        return string.Empty;
    }
}

public static class ExtensionMethods
{
    public static string CharSeparatedSort(this string self, char separator = ',')
    {
        return self.SplitAndTrim(separator).SortAndJoin(separator);
    }

    public static bool GetBoolean(this string self, bool defaultValue)
    {
        return self.GetBoolean().GetValueOrDefault(defaultValue);
    }

    public static bool? GetBoolean(this string self)
    {
        if (self.IsBoolean() && bool.TryParse(self, out bool result))
        {
            return result;
        }

        if (self.IsByte() && byte.TryParse(self, out byte result2))
        {
            return Convert.ToBoolean(result2);
        }

        if (self.IsDouble() && double.TryParse(self, out double result3))
        {
            return Convert.ToBoolean(result3);
        }

        return null;
    }

    public static DateTime? GetDateTime(this string self)
    {
        if (self.IsDateTime() && DateTime.TryParse(self, out DateTime result))
        {
            return result;
        }

        if (self.IsDouble() && double.TryParse(self, out double result2))
        {
            return DateTime.FromOADate(result2);
        }

        if (self.IsLong() && long.TryParse(self, out long result3))
        {
            return DateTime.FromOADate(result3);
        }

        return null;
    }

    public static bool IsBoolean(this string self)
    {
        if (!string.IsNullOrWhiteSpace(self) && bool.TryParse(self, out bool _))
        {
            return true;
        }

        return false;
    }

    public static bool IsByte(this string self)
    {
        if (!string.IsNullOrWhiteSpace(self) && byte.TryParse(self, out byte _))
        {
            return true;
        }

        return false;
    }

    public static bool IsDateTime(this string self)
    {
        if (!string.IsNullOrWhiteSpace(self) && DateTime.TryParse(self, out DateTime _))
        {
            return true;
        }

        return false;
    }

    public static bool IsDouble(this string self)
    {
        if (!string.IsNullOrWhiteSpace(self) && double.TryParse(self, out double _))
        {
            return true;
        }

        return false;
    }

    public static bool IsInteger(this string self)
    {
        if (!string.IsNullOrWhiteSpace(self) && int.TryParse(self, out int _))
        {
            return true;
        }

        return false;
    }

    public static bool IsLong(this string self)
    {
        if (!string.IsNullOrWhiteSpace(self) && long.TryParse(self, out long _))
        {
            return true;
        }

        return false;
    }

    public static bool IsNullOrEmpty<TSource>(this IEnumerable<TSource> enumerable)
    {
        if (enumerable != null && enumerable.Count() > 0)
        {
            return false;
        }

        return true;
    }

    public static bool IsNumber(this string self)
    {
        if (!self.IsDouble())
        {
            return self.IsLong();
        }

        return true;
    }

    public static string RemoveDoubleSpace(this string self)
    {
        string text = self.Validate();
        while (text.Contains("  "))
        {
            text = text.Replace("  ", " ");
        }

        return text;
    }

    public static string SortAndJoin(this IEnumerable<string> self, char separator = ',')
    {
        StringBuilder stringBuilder = new StringBuilder();
        if (self != null && self.Count() > 0)
        {
            List<string> list = [];
            foreach (string item in self)
            {
                list.Add(item.Trim().Trim(separator));
            }

            list.Sort();
            foreach (string item2 in list)
            {
                stringBuilder.Append(item2 + separator);
            }
        }

        return stringBuilder.ToString().Trim().Trim(separator);
    }

    public static IEnumerable<string> SplitAndTrim(this string self, char separator = ',')
    {
        List<string> list = [];
        if (!string.IsNullOrWhiteSpace(self))
        {
            string[] array = self.Split(separator);
            foreach (string text in array)
            {
                list.Add(text.Trim().Trim(separator));
            }
        }

        return list;
    }

    public static string ToSingleLine(this string self)
    {
        return self.Validate().Replace("\r\n", " ").Replace("\r", " ")
            .Replace("\n", " ")
            .RemoveDoubleSpace();
    }

    public static string Validate(this string self, params char[] trimChars)
    {
        string result = string.Empty;
        if (!string.IsNullOrWhiteSpace(self))
        {
            result = ((trimChars == null || trimChars.Length == 0) ? self.Trim() : self.Trim(trimChars));
        }

        return result;
    }
}