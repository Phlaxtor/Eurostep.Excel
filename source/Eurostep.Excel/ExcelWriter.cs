using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    public sealed class ExcelWriter : IDisposable
    {
        private readonly BorderStyleCache _borderStyleCache;
        private readonly CellStyleCache _cellStyleCache;
        private readonly FillStyleCache _fillStyleCache;
        private readonly FontStyleCache _fontStyleCache;
        private readonly NumberingFormatStyleCache _numberFormatStyleCache;
        private readonly Dictionary<string, ExcelWriterData> _reportCache = [];
        private readonly Stylesheet? _stylesheet;
        private readonly WorkbookPart? _workbookPart;
        private ExcelWriterData? _currentData;
        private WorksheetPart? _currentSheet;
        private SheetData? _currentSheetData;
        private bool _isDisposed = false;
        private uint _sheetCount;
        private SpreadsheetDocument? _spreasheet;
        private uint _tableCount;

        public ExcelWriter(Stream stream)
        {
            _spreasheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            _workbookPart = GetWorkbookPart();
            _stylesheet = GetStylesheet();
            _sheetCount = GetLastSheetId();
            _tableCount = 0;
            _borderStyleCache = new BorderStyleCache(this);
            _cellStyleCache = new CellStyleCache(this);
            _fillStyleCache = new FillStyleCache(this);
            _fontStyleCache = new FontStyleCache(this);
            _numberFormatStyleCache = new NumberingFormatStyleCache(this);
        }

        internal BorderStyleCache BorderStyles => _borderStyleCache;
        internal CellStyleCache CellStyles => _cellStyleCache;
        internal ExcelWriterData CurrentData => _currentData ?? throw new ApplicationException();
        internal WorksheetPart CurrentSheet => _currentSheet ?? throw new ApplicationException();
        internal SheetData CurrentSheetData => _currentSheetData ?? throw new ApplicationException();
        internal FillStyleCache FillStyles => _fillStyleCache;
        internal FontStyleCache FontStyles => _fontStyleCache;
        internal NumberingFormatStyleCache NumberingFormats => _numberFormatStyleCache;
        internal SpreadsheetDocument Spreasheet => _spreasheet ?? throw new ApplicationException();
        internal Stylesheet Stylesheet => _stylesheet ?? throw new ApplicationException();
        internal WorkbookPart WorkbookPart => _workbookPart ?? throw new ApplicationException();

        public void AddDataValidationCustom(ColumnRange validationRange, string formula, string? errorText = null, string? errorTitle = null)
        {
            DataValidation validation = CreateDataValidation(validationRange);
            validation.Type = DataValidationValues.Custom;
            validation.Formula1 = new Formula1() { Text = formula };
            AddValidationError(validation, errorText, errorTitle);
        }

        public void AddDataValidationForType(ColumnRange validationRange, DataValidationValues type, string? errorText = null, string? errorTitle = null)
        {
            DataValidation validation = CreateDataValidation(validationRange);
            validation.Type = type;
            AddValidationError(validation, errorText, errorTitle);
        }

        public void AddDataValidationList(ColumnRange validationRange, ColumnRange formulaForList, string? errorText = null, string? errorTitle = null)
        {
            DataValidation validation = CreateDataValidation(validationRange);
            validation.Type = DataValidationValues.List;
            validation.AllowBlank = new BooleanValue(true);
            AddValidationError(validation, errorText, errorTitle);
            Formula1 formula = new Formula1 { Text = formulaForList };
            validation.Append(formula);
        }

        public void AddDataValidationWhole(ColumnRange validationRange, int min = int.MinValue, int max = int.MaxValue, string? errorText = null, string? errorTitle = null)
        {
            DataValidation validation = CreateDataValidation(validationRange);
            validation.Type = DataValidationValues.Whole;
            validation.AllowBlank = new BooleanValue(true);
            AddValidationError(validation, errorText, errorTitle);
            Formula1 formula1 = new Formula1 { Text = min.ToString(), };
            validation.Append(formula1);
            Formula2 formula2 = new Formula2 { Text = max.ToString() };
            validation.Append(formula2);
        }

        public HeaderBuilder AddHeaders()
        {
            return new HeaderBuilder(this, CurrentData.Name);
        }

        public void AddHeaders(params IPresentationColumn[] headers)
        {
            if (headers.Length == 0) throw new ArgumentException($"Provided value is empty", nameof(headers));
            _tableCount++;
            CurrentData.IncreaseRowNo();
            CurrentData.SetHeaders(headers);
            CurrentData.StartTable(_tableCount);
            CellRef startCell = CurrentData.GetCurrentCell();
            NewColumnsData(startCell, headers);
            Spreasheet.Save();
        }

        public void AddMandatoryCellCheck(ColumnRange range, string condition, GeneralColor color)
        {
            IEnumerable<ConditionalFormatting> previous = CurrentSheet.Worksheet.ChildElements.OfType<ConditionalFormatting>();
            int count = previous.Count();
            int priority = count + 1;
            uint dxfId = CreateDifferentialColorFillFormat(bgColor: color);
            AddConditionalFormatting(range, ConditionalFormatValues.Expression, condition, dxfId, priority);
        }

        public uint AddNewRow(params string?[] values)
        {
            return AddNewRow(DefaultCellValue.Get(values));
        }

        public uint AddNewRow(params ICellValue[] values)
        {
            CurrentData.IncreaseRowNo();
            CellRef startCell = CurrentData.GetCurrentCell();
            WriteNewRowValues(startCell, values);
            return CurrentData.RowEnd;
        }

        public void AddNewTab(string name, params IPresentationColumn[] headers)
        {
            AddNewTab(name);
            Spreasheet.Save();
            if (headers.Length > 0) AddHeaders(headers);
        }

        public uint AddVerticalColumn(string header, CellStyleValue? style, params string?[] values)
        {
            CurrentData.IncreaseRowNo();
            ICellValue[] cells = new ICellValue[values.Length + 1];
            cells[0] = new DefaultCellValue(header, style);
            for (int i = 0; i < values.Length; i++)
            {
                cells[i + 1] = new DefaultCellValue(values[i]);
            }
            CellRef startCell = CurrentData.GetCurrentCell();
            WriteNewRowValues(startCell, cells);
            return CurrentData.RowEnd;
        }

        public void Close()
        {
            if (_isDisposed) return;
            Spreasheet.Save();
            Spreasheet.Dispose();
        }

        public void CloseTab(string name, bool addTable)
        {
            CurrentData.EndTable();
            CurrentSheet.Worksheet.Save();
            if (addTable)
            {
                AddTable(CurrentData.GetTableArea(), CurrentData.TableId);
            }
            else
            {
                AddFilter(CurrentData.GetTableArea());
            }

            CurrentSheet.Worksheet.Save();
            _reportCache.Remove(name);
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            _reportCache.Clear();
            _currentData = null;
            _currentSheet = null;
            _currentSheetData = null;
            _spreasheet = null;
            _isDisposed = true;
        }

        public void SetCurrentTab(string name)
        {
            if (_currentData?.Name == name) return;
            _currentData = _reportCache[name];
            _currentSheet = GetWorksheetPartBySheetName(_currentData.SheetName);
            _currentSheetData = _currentSheet.Worksheet.GetFirstChild<SheetData>();
        }

        internal uint AddBorder(Border value)
        {
            if (Stylesheet.Borders == null)
            {
                Stylesheet.Borders = new Borders();
            }

            uint count = Stylesheet.Borders.Count ?? 0;
            Stylesheet.Borders.Append(value);
            count++;
            Stylesheet.Borders.Count = count;
            Stylesheet.Save();
            return count;
        }

        internal uint AddCellFormat(CellFormat value)
        {
            if (Stylesheet.CellFormats == null)
            {
                Stylesheet.CellFormats = new CellFormats();
            }

            uint count = Stylesheet.CellFormats.Count ?? 0;
            Stylesheet.CellFormats.Append(value);
            count++;
            Stylesheet.CellFormats.Count = count;
            Stylesheet.Save();
            return count;
        }

        internal uint AddFill(Fill value)
        {
            if (Stylesheet.Fills == null)
            {
                Stylesheet.Fills = new Fills();
            }

            uint count = Stylesheet.Fills.Count ?? 0;
            Stylesheet.Fills.Append(value);
            count++;
            Stylesheet.Fills.Count = count;
            Stylesheet.Save();
            return count;
        }

        internal uint AddFont(Font value)
        {
            if (Stylesheet.Fonts == null)
            {
                Stylesheet.Fonts = new Fonts();
            }

            uint count = Stylesheet.Fonts.Count ?? 0;
            Stylesheet.Fonts.Append(value);
            count++;
            Stylesheet.Fonts.Count = count;
            Stylesheet.Save();
            return count;
        }

        internal uint AddNumberingFormat(NumberingFormat value)
        {
            if (Stylesheet.NumberingFormats == null)
            {
                Stylesheet.NumberingFormats = new NumberingFormats();
            }

            uint count = Stylesheet.NumberingFormats.Count ?? 0;
            Stylesheet.NumberingFormats.Append(value);
            count++;
            Stylesheet.NumberingFormats.Count = count;
            Stylesheet.Save();
            return count;
        }

        internal BorderStyleValue CreateBorder(IBorderStyle value)
        {
            return CreateBorder(value.TopBorder, value.RightBorder, value.BottomBorder, value.LeftBorder);
        }

        internal BorderStyleValue CreateBorder(BorderPart? top = null, BorderPart? right = null, BorderPart? bottom = null, BorderPart? left = null)
        {
            Border value = new Border();
            if (top.HasValue) value.TopBorder = top.Value.ToBorder<TopBorder>();
            if (right.HasValue) value.RightBorder = right.Value.ToBorder<RightBorder>();
            if (bottom.HasValue) value.BottomBorder = bottom.Value.ToBorder<BottomBorder>();
            if (left.HasValue) value.LeftBorder = left.Value.ToBorder<LeftBorder>();
            uint result = AddBorder(value);
            return new BorderStyleValue(result, top, right, bottom, left);
        }

        internal uint CreateCellFormat(uint? numFmtId = null, uint? xfId = null, Alignment? alignment = null, uint? fontId = null, uint? borderId = null, uint? fillId = null, Protection? protection = null, bool? pivotButton = null, bool? quotePrefix = null)
        {
            CellFormat value = new CellFormat();
            if (pivotButton.HasValue) value.PivotButton = pivotButton.Value;
            if (quotePrefix.HasValue) value.QuotePrefix = quotePrefix.Value;
            if (xfId.HasValue) value.FormatId = xfId;
            if (protection != null)
            {
                value.Protection = protection;
                value.ApplyProtection = true;
            }
            if (alignment != null)
            {
                value.Alignment = alignment;
                value.ApplyAlignment = true;
            }
            if (borderId.HasValue)
            {
                value.BorderId = borderId.Value;
                value.ApplyBorder = true;
            }
            if (fontId.HasValue)
            {
                value.FontId = fontId.Value;
                value.ApplyFont = true;
            }
            if (fillId.HasValue)
            {
                value.FillId = fillId.Value;
                value.ApplyFill = true;
            }
            if (numFmtId.HasValue)
            {
                value.NumberFormatId = numFmtId.Value;
                value.ApplyNumberFormat = true;

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
            uint result = AddCellFormat(value);
            return result;
        }

        internal FillStyleValue CreateFill(IFillStyle value)
        {
            return CreateFill(value.PatternType, value.ForegroundColor, value.BackgroundColor, value.GradientType, value.GradientDegree, value.GradientTop, value.GradientBottom, value.GradientRight, value.GradientLeft);
        }

        internal FillStyleValue CreateFill(PatternValues? patternType = null, GeneralColor? fgColor = null, GeneralColor? bgColor = null, GradientValues? gradientType = null, double degree = 0, double top = 0, double bottom = 0, double right = 0, double left = 0)
        {
            Fill value = new Fill
            {
                PatternFill = new PatternFill()
            };
            if (fgColor.HasValue) value.PatternFill.ForegroundColor = fgColor.Value.ToSpreadsheetColor<ForegroundColor>();
            if (bgColor.HasValue) value.PatternFill.BackgroundColor = bgColor.Value.ToSpreadsheetColor<BackgroundColor>();
            value.PatternFill.PatternType = patternType;
            if (gradientType.HasValue)
            {
                value.GradientFill = new GradientFill
                {
                    Type = gradientType.Value,
                    Degree = degree,
                    Top = top,
                    Bottom = bottom,
                    Right = right,
                    Left = left
                };
            }
            uint result = AddFill(value);
            return new FillStyleValue(result, patternType, fgColor, bgColor, gradientType, degree, top, bottom, right, left);
        }

        internal FontStyleValue CreateFont(IFontStyle value)
        {
            return CreateFont(value.Name, value.Size, value.Bold, value.Italic, value.Underline, value.Color, value.VerticalTextAlignment, value.Strike, value.Condense, value.Extend, value.Shadow);
        }

        internal FontStyleValue CreateFont(string? name = "Calibri", double? sz = 11, bool? b = null, bool? i = null, UnderlineValues? u = null, GeneralColor? color = null, VerticalAlignmentRunValues? vertAlig = null, bool? strike = null, bool? condense = null, bool? extend = null, bool? shadow = null)
        {
            Font value = new Font();
            if (string.IsNullOrEmpty(name) == false) value.FontName = new FontName { Val = name };
            if (sz.HasValue) value.FontSize = new FontSize { Val = sz.Value };
            if (b.HasValue) value.Bold = new Bold { Val = b.Value };
            if (i.HasValue) value.Italic = new Italic { Val = i.Value };
            if (u.HasValue) value.Underline = new Underline { Val = u.Value };
            if (strike.HasValue) value.Strike = new Strike { Val = strike.Value };
            if (condense.HasValue) value.Condense = new Condense { Val = condense.Value };
            if (extend.HasValue) value.Extend = new Extend { Val = extend.Value };
            if (shadow.HasValue) value.Shadow = new Shadow { Val = shadow.Value };
            if (vertAlig.HasValue) value.VerticalTextAlignment = new VerticalTextAlignment { Val = vertAlig.Value };
            if (color.HasValue) value.Color = color.Value.ToSpreadsheetColor<Color>();
            uint result = AddFont(value);
            return new FontStyleValue(result, name, sz, b, i, u, color, vertAlig, strike, condense, extend, shadow);
        }

        internal NumberingFormatStyleValue CreateNumberingFormat(INumberingFormatStyle value)
        {
            return CreateNumberingFormat(value.NumberFormatId, value.FormatCode);
        }

        internal NumberingFormatStyleValue CreateNumberingFormat(uint? numFmtId = 0, string? formatCode = null)
        {
            NumberingFormat value = new NumberingFormat
            {
                FormatCode = formatCode,
                NumberFormatId = numFmtId
            };
            uint result = AddNumberingFormat(value);
            return new NumberingFormatStyleValue(result, numFmtId, formatCode);
        }

        internal Alignment GetAlignment(HorizontalAlignmentValues horizontal, VerticalAlignmentValues vertical, uint indent = 0, int relativeIndent = 0, bool shrinkToFit = false, bool wrapText = false, uint textRotation = 0, string? mergeCell = null, uint readingOrder = 0, bool justifyLastLine = false)
        {
            Alignment result = new Alignment
            {
                Horizontal = horizontal,
                Indent = indent,
                JustifyLastLine = justifyLastLine,
                MergeCell = mergeCell,
                ReadingOrder = readingOrder,
                RelativeIndent = relativeIndent,
                ShrinkToFit = shrinkToFit,
                TextRotation = textRotation,
                Vertical = vertical,
                WrapText = wrapText
            };
            return result;
        }

        internal Protection GetProtection(bool hidden, bool locked)
        {
            Protection result = new Protection
            {
                Hidden = hidden,
                Locked = locked
            };
            return result;
        }

        internal CellStyleValue NewCellStyle(NumberingFormatStyleValue? numberingFormat, uint? formatId, Alignment? alignment, FontStyleValue? font, BorderStyleValue? border, FillStyleValue? fill, Protection? protection, bool? pivotButton, bool? quotePrefix)
        {
            //DocumentFormat.OpenXml.Spreadsheet.CellFormat
            uint? numberFormatId = numberingFormat.HasValue ? numberingFormat.Value.NumberFormatId : null;
            uint cellFormat = CreateCellFormat(numberFormatId, formatId, alignment, font, border, fill, protection, pivotButton, quotePrefix);
            CellStyleValue value = new CellStyleValue(cellFormat, formatId, alignment, border, fill, font, numberingFormat, pivotButton, protection, quotePrefix);
            return value;
        }

        private void AddConditionalFormatting(ColumnRange range, ConditionalFormatValues type, string condition, uint dxfId, int priority)
        {
            Formula formula = new Formula { Text = condition };
            ConditionalFormattingRule cfRule = new ConditionalFormattingRule
            {
                Type = type,
                FormatId = dxfId,
                Priority = priority,
            };
            cfRule.AddChild(formula);
            ConditionalFormatting formatting = new ConditionalFormatting
            {
                SequenceOfReferences = new ListValue<StringValue>() { InnerText = range.GetRelative() },
            };
            formatting.AddChild(cfRule);
            InsertWorksheetChildElement(formatting);
        }

        private void AddDataValidation(DataValidation dataValidation)
        {
            DataValidations? dataValidations = CurrentSheet.Worksheet.GetFirstChild<DataValidations>();
            if (dataValidations == null)
            {
                dataValidations = new DataValidations
                {
                    Count = 0
                };
                InsertWorksheetChildElement(dataValidations);
            }
            uint count = dataValidations.Count?.Value ?? 0;
            count++;
            dataValidations.Count = count;
            dataValidations.Append(dataValidation);
        }

        private uint AddDifferentialFormat(DifferentialFormat dxf)
        {
            if (Stylesheet.DifferentialFormats == null)
            {
                Stylesheet.DifferentialFormats = new DifferentialFormats();
            }
            uint count = Stylesheet.DifferentialFormats.Count ?? 0;
            Stylesheet.DifferentialFormats.Append(dxf);
            count++;
            Stylesheet.DifferentialFormats.Count = count;
            Stylesheet.Save();
            return count - 1;
        }

        private ExcelWriterData AddExcelWriterData(string name, uint sheetNo, uint columnStart = 1, uint rowStart = 1)
        {
            if (_reportCache.TryGetValue(name, out ExcelWriterData? data) == true) throw new ApplicationException();
            data = new ExcelWriterData(name, sheetNo, columnStart, rowStart);
            _reportCache[name] = data;
            return data;
        }

        private void AddFilter(CellArea area)
        {
            AutoFilter autoFilter = new AutoFilter
            {
                Reference = area.ToString(),
            };
            InsertWorksheetChildElement(autoFilter);
        }

        private void AddNewTab(string name)
        {
            _sheetCount++;
            ExcelWriterData data = AddExcelWriterData(name, _sheetCount);
            AddSheet(data.SheetName, data.SheetNo);
            SetCurrentTab(name);
        }

        private void AddSheet(string sheetName, uint sheetId)
        {
            Sheets sheets = Spreasheet.WorkbookPart?.Workbook?.GetFirstChild<Sheets>() ?? throw new ApplicationException();
            WorksheetPart worksheetPart = Spreasheet.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            Sheet sheet = new Sheet
            {
                Id = Spreasheet.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = sheetName
            };
            sheets.Append(sheet);
            worksheetPart.Worksheet.Save();
        }

        private void AddTable(CellArea area, uint id, string tableStyleName = "TableStyleMedium2")
        {
            if (area.HasRows)
            {
                string definitionId = $"rId{id}";
                TableDefinitionPart tableDefinitionPart = CurrentSheet.AddNewPart<TableDefinitionPart>(definitionId);
                tableDefinitionPart.Table = CreateTable(area, id, tableStyleName);
                IEnumerable<TableParts> tablePartsCollection = CurrentSheet.Worksheet.Elements<TableParts>();
                TableParts? tableParts = tablePartsCollection.FirstOrDefault();
                if (tableParts == null)
                {
                    tableParts = new TableParts { Count = 0 };
                    CurrentSheet.Worksheet.Append(tableParts);
                }

                uint count = tableParts.Count ?? 0;
                TablePart tablePart = new TablePart { Id = definitionId };
                tableParts.Count = count + 1;
                tableParts.Append(tablePart);
                CurrentSheet.Worksheet.Save();

                //var dxfs = WorkbookPart.WorkbookStylesPart.RootElement.Descendants<DifferentialFormats>().FirstOrDefault();
                //var differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)1U };
                //var differentialFormat1 = new DifferentialFormat();
                //var numberingFormat2 = new NumberingFormat() { NumberFormatId = (UInt32Value)10U, FormatCode = "\"$\"#,##0_);[Red]\\(\"$\"#,##0\\)" };
                //differentialFormat1.Append(numberingFormat2);
                //differentialFormats1.Append(differentialFormat1);
                //var tableStyles = new TableStyles { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleMedium9" };
                //if (dxfs == null)
                //{
                //    WorkbookPart.WorkbookStylesPart.RootElement.AddChild(differentialFormats1);
                //}
                //else
                //{
                //    WorkbookPart.WorkbookStylesPart.RootElement.ReplaceChild(differentialFormats1, dxfs);
                //}
            }
        }

        private void AddValidationError(DataValidation validation, string? errorText, string? errorTitle)
        {
            if (string.IsNullOrEmpty(errorText)) return;
            if (string.IsNullOrEmpty(errorTitle)) return;
            validation.ShowErrorMessage = true;
            validation.ErrorTitle = errorTitle;
            validation.Error = errorText;
        }

        private Column CreateColumn(ColumnId column, double columnWidth = 10, CellStyleValue? style = default)
        {
            return CreateColumn(column, column, columnWidth, style);
        }

        private Column CreateColumn(ColumnId startColumn, ColumnId endColumn, double columnWidth = 10, CellStyleValue? style = default)
        {
            columnWidth = columnWidth >= 10 ? columnWidth : 10;
            Column column = new Column
            {
                Min = startColumn.No,
                Max = endColumn.No,
                Width = columnWidth,
                CustomWidth = true,
                BestFit = true,
                Collapsed = false,
                Hidden = false,
                Style = (uint?)style
            };
            return column;
        }

        private DataValidation CreateDataValidation(ColumnRange validationRange)
        {
            DataValidation validation = new DataValidation()
            {
                AllowBlank = true,
                SequenceOfReferences = new ListValue<StringValue>() { InnerText = validationRange }
            };
            AddDataValidation(validation);
            return validation;
        }

        private uint CreateDifferentialColorFillFormat(PatternValues? patternType = null, GeneralColor? fgColor = null, GeneralColor? bgColor = null)
        {
            DifferentialFormat differentialFormat = new DifferentialFormat();
            Fill format = new Fill
            {
                PatternFill = new PatternFill { PatternType = patternType }
            };
            if (bgColor.HasValue) format.PatternFill.BackgroundColor = bgColor.Value.ToSpreadsheetColor<BackgroundColor>();
            if (fgColor.HasValue) format.PatternFill.ForegroundColor = fgColor.Value.ToSpreadsheetColor<ForegroundColor>();
            differentialFormat.AddChild(format);
            return AddDifferentialFormat(differentialFormat);
        }

        private uint CreateDifferentialNumberingFormat(uint numFmtId, string formatCode)
        {
            DifferentialFormat differentialFormat = new DifferentialFormat();
            NumberingFormat format = new NumberingFormat { NumberFormatId = numFmtId, FormatCode = formatCode };
            differentialFormat.Append(format);
            return AddDifferentialFormat(differentialFormat);
        }

        private Table CreateTable(CellArea area, uint id, string tableStyleName)
        {
            string name = $"Table{id}";
            Table table = new Table
            {
                Id = id,
                Name = name,
                DisplayName = name,
                Reference = area.ToString(),
                TotalsRowShown = false
            };

            AutoFilter autoFilter = new AutoFilter
            {
                Reference = area.ToString()
            };

            string sortConditionReference = $"{area.StartColumn}{area.StartRow + 1}:{area.GetLowerLeft()}";
            SortCondition sortCondition = new SortCondition
            {
                Reference = StringValue.ToString(sortConditionReference),
                Descending = BooleanValue.ToBoolean(true)
            };

            string sortStateReference = $"{area.StartColumn}{area.StartRow + 1}:{area.GetLowerRight()}";
            SortState sortState = new SortState
            {
                Reference = StringValue.ToString(sortStateReference)
            };
            sortState.Append(sortCondition);

            TableColumns tableColumns = new TableColumns { Count = area.TotalColumns };
            for (uint i = area.StartColumn; i <= area.EndColumn; i++)
            {
                CellRef header = area.Start.GetForColumn(i);
                string? colVal = GetCellValue(header);
                TableColumn tableColumn = new TableColumn { Id = i, Name = colVal };
                tableColumns.Append(tableColumn);
            }

            TableStyleInfo tableStyleInfo = new TableStyleInfo
            {
                Name = tableStyleName,
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = false
            };

            table.Append(autoFilter);
            table.Append(sortState);
            table.Append(tableColumns);
            table.Append(tableStyleInfo);
            return table;
        }

        private string? GetCellValue(CellRef cellRef)
        {
            IEnumerable<Row> rows = CurrentSheetData.Elements<Row>();
            Row? row = rows.FirstOrDefault(r => r?.RowIndex?.Value == cellRef.RowId);
            if (row == null) return null;
            string cellReference = cellRef.ToString();
            IEnumerable<Cell> cells = row.Elements<Cell>();
            Cell? cell = cells.FirstOrDefault(c => c?.CellReference?.Value == cellReference);
            if (cell == null) return null;
            return GetCellValue(cell);
        }

        private string? GetCellValue(Cell cell)
        {
            string? value = cell.CellValue?.Text;
            if (string.IsNullOrEmpty(value)) return null;

            // If the content of the first cell is stored as a shared string, get the text of the first cell
            // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                SharedStringTablePart shareStringPart = WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                SharedStringItem[] items = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
                return items[int.Parse(value)].InnerText;
            }

            return value;
        }

        private Columns GetColumns()
        {
            Columns? columns = CurrentSheet.Worksheet.GetFirstChild<Columns>();
            if (columns == null)
            {
                columns = new Columns();
                CurrentSheet.Worksheet.InsertAfter(columns, CurrentSheet.Worksheet.SheetFormatProperties);
            }
            return columns;
        }

        private Stylesheet GetDefaultStylesheet()
        {
            return new Stylesheet(
                new Fonts(
                    new Font(
                        new FontSize { Val = 11 },
                        new DocumentFormat.OpenXml.Office2010.Excel.Color { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName { Val = "Calibri" }))
                {
                    Count = 0
                },
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 }))
                {
                    Count = 1
                },
                new Borders(
                    new Border(
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()))
                {
                    Count = 0
                },
                new CellFormats(new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 })
                {
                    Count = 0
                });
        }

        private uint GetLastSheetId()
        {
            IEnumerable<Sheet> sheets = WorkbookPart.Workbook.Descendants<Sheet>();
            if (sheets == null) return 0;
            uint result = 0;
            foreach (Sheet sheet in sheets)
            {
                uint sheetId = sheet.SheetId ?? 0;
                if (sheetId > result) result = sheetId;
            }
            return result;
        }

        private Cell GetNewCell(CellRef cellRef, ICellValue value)
        {
            Cell result = new Cell
            {
                CellReference = StringValue.FromString(cellRef.ToString()),
                DataType = new EnumValue<CellValues>(value.DataType),
                StyleIndex = value.Style.HasValue ? new UInt32Value(value.Style.Value.Value) : default
            };
            if (string.IsNullOrEmpty(value.Value)) return result;
            switch (value.DataType)
            {
                case CellValues.Boolean:
                case CellValues.Number:
                case CellValues.String:
                case CellValues.Date:
                    result.CellValue = new CellValue(value.Value);
                    break;

                case CellValues.InlineString:
                    Text text = new Text { Text = value.Value };
                    InlineString inlineString = new InlineString();
                    inlineString.AppendChild(text);
                    result.AppendChild(inlineString);
                    break;

                case CellValues.Error:
                    throw new ApplicationException($"Error in cell '{cellRef}'");
                case CellValues.SharedString:
                    throw new NotImplementedException($"Not implemented support for shared strings");
                default:
                    throw new NotSupportedException($"Not supported value: '{value.DataType}'");
            }
            return result;
        }

        private List<Type> GetPossiblePredecessors(OpenXmlElement child, Type[] sequence)
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

        private List<Type> GetPossibleSuccessors(OpenXmlElement child, Type[] sequence)
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

        private Stylesheet GetStylesheet()
        {
            WorkbookStylesPart? stylesPart = WorkbookPart.WorkbookStylesPart;
            if (stylesPart == null)
            {
                stylesPart = WorkbookPart.AddNewPart<WorkbookStylesPart>();
            }

            if (stylesPart.Stylesheet != null)
            {
                return stylesPart.Stylesheet;
            }

            Stylesheet stylesheet = GetDefaultStylesheet();
            stylesPart.Stylesheet = stylesheet;
            stylesheet.Save();
            return stylesheet;
        }

        private WorkbookPart GetWorkbookPart()
        {
            WorkbookPart workbookPart = Spreasheet.WorkbookPart ?? Spreasheet.AddWorkbookPart();

            if (workbookPart.Workbook == null)
            {
                workbookPart.Workbook = new Workbook();
                workbookPart.Workbook.AppendChild(new Sheets());
            }
            Spreasheet.Save();
            return workbookPart;
        }

        private WorksheetPart GetWorksheetPartBySheetId(string sheetId)
        {
            return (WorksheetPart)WorkbookPart.GetPartById(sheetId);
        }

        private WorksheetPart? GetWorksheetPartBySheetName(string sheetName)
        {
            IEnumerable<Sheet> sheets = WorkbookPart.Workbook.Descendants<Sheet>();
            if (sheets != null)
            {
                foreach (Sheet sheet in sheets)
                {
                    if (string.Equals(sheetName, sheet.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        return GetWorksheetPartBySheetId(sheet.Id);
                    }
                }
            }

            return null;
        }

        private void InsertWorksheetChildElement(OpenXmlElement child)
        {
            int sheetDataPosition = 5;
            // NB: Worksheet children must be appended in the correct order (matching the order in the sequence of CT_Worksheet in sml.xsd)
            // we can assumme that SheetData is always present
            Type[] sequence = new Type[]
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

            bool isBeforeSheetData = false;
            for (int i = 0; i < sheetDataPosition; i++)
            {
                if (child.GetType().Name == sequence[i].Name)
                {
                    isBeforeSheetData = true;
                    break;
                }
            }

            if (isBeforeSheetData)
            {
                InsertWorksheetChildElementBefore(child, sequence);
            }
            else
            {
                InsertWorksheetChildElementAfter(child, sequence);
            }
        }

        private void InsertWorksheetChildElementAfter(OpenXmlElement child, Type[] sequence)
        {
            List<Type> possiblePredecessors = GetPossiblePredecessors(child, sequence); new List<Type>();
            foreach (OpenXmlElement? element in CurrentSheet.Worksheet.ChildElements.Reverse())
            {
                if (possiblePredecessors.Contains(element.GetType()))
                {
                    CurrentSheet.Worksheet.InsertAfter(child, element);
                    return;
                }
            }
            CurrentSheet.Worksheet.AppendChild(child);
        }

        private void InsertWorksheetChildElementBefore(OpenXmlElement child, Type[] sequence)
        {
            List<Type> possiblePredecessors = GetPossibleSuccessors(child, sequence); new List<Type>();
            foreach (OpenXmlElement element in CurrentSheet.Worksheet.ChildElements)
            {
                if (possiblePredecessors.Contains(element.GetType()))
                {
                    CurrentSheet.Worksheet.InsertBefore(child, element);
                    return;
                }
            }
            CurrentSheet.Worksheet.AppendChild(child);
        }

        private void NewColumnsData(CellRef startCell, params IPresentationColumn[] headers)
        {
            ICellValue[] cellValues = new ICellValue[headers.Length];
            Columns columns = GetColumns();
            for (uint i = 0; i < headers.Length; i++)
            {
                CellRef headerCell = startCell.GetForColumn(i);
                IPresentationColumn header = headers[i];
                string columnName = string.IsNullOrEmpty(header.DisplayName) ? $"Column {headerCell.Column}" : header.DisplayName;
                cellValues[i] = new DefaultCellValue(columnName, header.HeaderStyle);
                Column column = CreateColumn(headerCell.Column, header.Width, header.ColumnStyle);
                columns.Append(column);
            }

            WriteNewRowValues(startCell, cellValues);
        }

        private void WriteNewRowValues(CellRef startCell, params ICellValue[] cellValues)
        {
            Row row = new Row
            {
                RowIndex = startCell.RowId,
            };
            CurrentSheetData.Append(row);
            uint startColumn = startCell.Column.No;
            for (uint i = 0; i < cellValues.Length; i++)
            {
                ICellValue cellValue = cellValues[i];
                if (string.IsNullOrEmpty(cellValue.Value)) continue;
                uint columnIndex = startColumn + i;
                CellRef cellReference = startCell.GetForColumn(columnIndex);
                Cell cell = GetNewCell(cellReference, cellValue);
                row.Append(cell);
            }
        }
    }
}