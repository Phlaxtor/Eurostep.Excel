using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using System.Reflection;

namespace FMA.Digimat.Toolbox.Excel.Model
{
    public abstract class ExcelRow
    {
        public uint Uid { get; set; }
        public uint RowId { get; private set; }
        public string DetailsForLogging { get; private set; }
        public abstract Dictionary<string, string> HeadingsWithColumnNames { get; }
        public abstract string SheetName { get; }
        public abstract uint FirstDataRow { get; }
        public abstract uint HeaderRow { get; }

        internal virtual void ReadRow(KeyValuePair<uint, Dictionary<string, string>> row)
        {
            RowId = row.Key;
            DetailsForLogging = $"sheet: {SheetName}; row: {RowId}";
            Type objtype = GetType();
            foreach (PropertyInfo p in objtype.GetProperties())
            {
                object? fieldNameAttribute = p.GetCustomAttributes(false).FirstOrDefault(z => z is ExcelColumnAttribute);
                if (fieldNameAttribute != null && HeadingsWithColumnNames.ContainsKey(((ExcelColumnAttribute)fieldNameAttribute).Heading))
                {
                    string key = HeadingsWithColumnNames[((ExcelColumnAttribute)fieldNameAttribute).Heading];

                    if (p.PropertyType == typeof(string))
                    {
                        p.SetValue(this, GetStringValue(row.Value, key));
                    }
                    else if (p.PropertyType == typeof(DateTime?))
                    {
                        p.SetValue(this, GetDateTimeValue(row.Value, key));
                    }
                    else if (p.PropertyType == typeof(int?))
                    {
                        p.SetValue(this, GetIntValue(row.Value, key));
                    }
                    else if (p.PropertyType == typeof(double?))
                    {
                        p.SetValue(this, GetDoubleValue(row.Value, key));
                    }
                    else
                    {
                        throw new NotImplementedException("Type " + p.PropertyType + " is not supported.");
                    }
                }
            }
        }

        private string? GetStringValue(Dictionary<string, string> row, string key)
        {
            if (key != null && row.ContainsKey(key))
            {
                return (row[key]);
            }
            return null;
        }

        private int? GetIntValue(Dictionary<string, string> row, string key)
        {
            string value = GetStringValue(row, key);
            if (!string.IsNullOrWhiteSpace(value) && int.TryParse(value, out int ret))
            {
                return ret;
            }
            return null;
        }

        private double? GetDoubleValue(Dictionary<string, string> row, string key)
        {
            string value = GetStringValue(row, key);
            if (!string.IsNullOrWhiteSpace(value) && double.TryParse(value, out double ret))
            {
                return ret;
            }
            return null;
        }

        private DateTime? GetDateTimeValue(Dictionary<string, string> row, string key)
        {
            DateTime? parsedDate = null;
            string value = GetStringValue(row, key);
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

        public virtual uint? GetRuntimeCellStyle(IExcelStyleCache styleCache, PropertyInfo property, string column, uint row)
        {
            return null;
        }

        internal static void WriteData<T>(Stream stream, ReportHeader header, IEnumerable<ExcelRow> excelRows, string sheetName) where T : ExcelRow
        {
            ColumnInfo[] columnInfos = GetColumnInfos<T>();
            ColumnInfo firstColumn = columnInfos.First();
            ColumnInfo lastColumn = columnInfos.Last();
            using SpreadsheetDocument doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = doc.GetWorkbookPart();
            WorksheetPart worksheetPart = workbookPart.GetWorksheetPart();
            workbookPart.AddSheet(sheetName);
            ExcelStyleCache stylesheetCache = new ExcelStyleCache(doc);
            uint rowNo = 0;
            foreach (ColumnInfo columnInfo in columnInfos)
            {
                DocumentFormat.OpenXml.Spreadsheet.Column column = worksheetPart.AddColumn(columnInfo.ColumnNo, columnInfo.Width);
                if (columnInfo.ColumnStyle is not null)
                {
                    column.Style = stylesheetCache.GetCellStyle(columnInfo.ColumnStyle);
                }
            }

            uint? titleStyle = stylesheetCache.GetCellStyle<ReportTitleStyle>();
            TitleHeaderFontStyle ths = new TitleHeaderFontStyle();
            TitleHeaderTextFontStyle thts = new TitleHeaderTextFontStyle();

            rowNo++;
            workbookPart.AddRowWithCell("A", rowNo, "DIGIMAT Material Delivery Report", 46, titleStyle);
            rowNo += 2;
            workbookPart.AddRowWithCell("A", rowNo, CellText.New("Project: ", ths), CellText.New(header.Project, thts));
            rowNo++;
            workbookPart.AddRowWithCell("A", rowNo, CellText.New("Supplier: ", ths), CellText.New(header.Supplier, thts));
            rowNo++;
            workbookPart.AddRowWithCell("A", rowNo, CellText.New("Work package order: ", ths), CellText.New(header.Workpackage, thts));
            rowNo++;
            workbookPart.AddRowWithCell("A", rowNo, CellText.New("Report generated on: ", ths), CellText.New(header.CreatedOn.ToShortDateString(), thts));
            rowNo++;
            workbookPart.AddRowWithCell("A", rowNo, CellText.New("Section/Category: ", ths), CellText.New(header.Section, thts));
            rowNo += 3;

            uint groupHeaderRow = rowNo;
            uint lastMergeColumn = 0;
            foreach (ColumnInfo columnInfo in columnInfos)
            {
                if (string.IsNullOrEmpty(columnInfo.MergeCellTitle))
                {
                    continue;
                }
                if (lastMergeColumn >= columnInfo.ColumnNo)
                {
                    throw new ApplicationException($"Overlapping merge cell, {columnInfo.Column} is starting before previous merge cell is ended");
                }
                if (columnInfo.MergeCellCount > 0)
                {
                    lastMergeColumn = columnInfo.ColumnNo + columnInfo.MergeCellCount;
                    worksheetPart.MergeRowCells(rowNo, columnInfo.ColumnNo, lastMergeColumn);
                }
                uint? mergeStyle = stylesheetCache.GetCellStyle(columnInfo.MergeCellStyle);
                worksheetPart.WriteStringValueInCell(columnInfo.ColumnNo, rowNo, columnInfo.MergeCellTitle, mergeStyle);
            }

            rowNo++;
            uint headerRow = rowNo;
            foreach (ColumnInfo columnInfo in columnInfos)
            {
                uint? headerStyle = stylesheetCache.GetCellStyle(columnInfo.HeaderStyle);
                worksheetPart.WriteStringValueInCell(columnInfo.ColumnNo, rowNo, columnInfo.Heading, headerStyle);
            }

            uint firstContentRow = rowNo + 1;
            foreach (T excelRow in excelRows)
            {
                rowNo++;
                excelRow.RowId = rowNo;
                foreach (ColumnInfo columnInfo in columnInfos)
                {
                    object? value = columnInfo.Property.GetValue(excelRow);
                    string excelValue = value?.ToString() ?? string.Empty;
                    uint? styleIndex = excelRow.GetRuntimeCellStyle(stylesheetCache, columnInfo.Property, columnInfo.Column, rowNo);
                    if (styleIndex.HasValue == false)
                    {
                        styleIndex = stylesheetCache.GetCellStyle(columnInfo.CellStyle);
                    }
                    worksheetPart.WriteStringValueInCell(columnInfo.ColumnNo, rowNo, excelValue, styleIndex);
                }
            }

            worksheetPart.FreezePanes("I", firstContentRow);
            worksheetPart.AddFilter(firstColumn.Column, headerRow, lastColumn.Column, rowNo);
            worksheetPart.SetSheetDimension(lastColumn.Column, rowNo);
            workbookPart.Workbook.Save();
        }

        private static ColumnInfo[] GetColumnInfos<T>() where T : ExcelRow
        {
            Dictionary<string, ColumnInfo> result = [];
            Type rowType = typeof(T);
            foreach (PropertyInfo property in rowType.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
            {
                ExcelColumnAttribute? columnAttr = property.GetCustomAttribute<ExcelColumnAttribute>();
                if (columnAttr is null)
                {
                    continue;
                }

                uint columnNo = columnAttr.Column.GetColumnIndexFromName();
                ExcelCellStyleBaseAttribute? cellStyle = property.GetCustomAttribute<ExcelCellStyleBaseAttribute>();
                ExcelColumnStyleBaseAttribute? columnStyle = property.GetCustomAttribute<ExcelColumnStyleBaseAttribute>();
                ExcelHeaderStyleBaseAttribute? headerStyle = property.GetCustomAttribute<ExcelHeaderStyleBaseAttribute>();
                ExcelMergeColumnsBaseAttribute? mergeStyle = property.GetCustomAttribute<ExcelMergeColumnsBaseAttribute>();
                ColumnInfo columnInfo = new ColumnInfo
                {
                    CellStyle = cellStyle?.GetStyle(),
                    Column = columnAttr.Column,
                    ColumnNo = columnNo,
                    ColumnStyle = columnStyle?.GetStyle(),
                    HeaderStyle = headerStyle?.GetStyle(),
                    Heading = columnAttr.Heading,
                    MergeCellCount = mergeStyle?.Count ?? 0,
                    MergeCellStyle = mergeStyle?.GetStyle(),
                    MergeCellTitle = mergeStyle?.Title,
                    Optional = columnAttr.Optional,
                    Property = property,
                    Width = columnAttr.Width,
                };

                if (result.TryGetValue(columnInfo.Column, out ColumnInfo? duplicated))
                {
                    throw new InvalidDataException($"Duplicate column '{columnInfo.Column}', column with header '{columnInfo.Heading}' exist already with heading '{duplicated.Heading}'");
                }

                result.Add(columnInfo.Column, columnInfo);
            }
            return result.Values.OrderBy(c => c.ColumnNo).ToArray();
        }
    }
}