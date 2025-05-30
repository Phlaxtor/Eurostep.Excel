﻿using DocumentFormat.OpenXml.Spreadsheet;

namespace Eurostep.Excel
{
    public interface ICellValue
    {
        CellValues DataType { get; }
        CellStyleValue? Style { get; }
        string? Value { get; }
    }
}