﻿namespace Eurostep.Excel
{
    public interface IPresentationColumn
    {
        CellStyleValue? ColumnStyle { get; }
        string DisplayName { get; }
        CellStyleValue? HeaderStyle { get; }
        double Width { get; }
    }
}