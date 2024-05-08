using System.Diagnostics.CodeAnalysis;

namespace Eurostep.Excel
{
    public readonly struct CellArea
    {
        public CellArea(string startColumn, uint startRow, string endColumn, uint endRow)
        {
            StartColumn = startColumn;
            StartRow = startRow;
            EndColumn = endColumn;
            EndRow = endRow;
            Start = new CellRef(startColumn, startRow);
            End = new CellRef(endColumn, endRow);
        }

        public CellArea(uint startColumn, uint startRow, uint endColumn, uint endRow)
        {
            StartColumn = startColumn;
            StartRow = startRow;
            EndColumn = endColumn;
            EndRow = endRow;
            Start = new CellRef(startColumn, startRow);
            End = new CellRef(endColumn, endRow);
        }

        public CellArea(ColumnId startColumn, uint startRow, ColumnId endColumn, uint endRow)
        {
            StartColumn = startColumn;
            StartRow = startRow;
            EndColumn = endColumn;
            EndRow = endRow;
            Start = new CellRef(startColumn, startRow);
            End = new CellRef(endColumn, endRow);
        }

        public CellArea(CellRef upperLeft, CellRef lowerRight)
        {
            StartColumn = upperLeft.Column;
            StartRow = upperLeft.RowId;
            EndColumn = lowerRight.Column;
            EndRow = lowerRight.RowId;
            Start = upperLeft;
            End = lowerRight;
        }

        public CellRef End { get; }
        public ColumnId EndColumn { get; }
        public uint EndRow { get; }
        public bool HasRows => StartRow < EndRow;
        public CellRef Start { get; }
        public ColumnId StartColumn { get; }
        public uint StartRow { get; }
        public uint TotalColumns => EndColumn - StartColumn + 1;
        public uint TotalRows => EndRow - StartRow + 1;

        public static implicit operator string(CellArea c) => c.ToString();

        public override bool Equals([NotNullWhen(true)] object? obj)
        {
            if (obj is not CellArea other) return false;
            if (other.StartColumn != StartColumn) return false;
            if (other.StartRow != StartRow) return false;
            if (other.EndColumn != EndColumn) return false;
            if (other.EndRow != EndRow) return false;
            return true;
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        public CellRef GetLowerLeft()
        {
            return new CellRef(StartColumn, EndRow);
        }

        public CellRef GetLowerRight()
        {
            return new CellRef(EndColumn, EndRow);
        }

        public CellRef GetUpperLeft()
        {
            return new CellRef(StartColumn, StartRow);
        }

        public CellRef GetUpperRight()
        {
            return new CellRef(EndColumn, StartRow);
        }

        public override string ToString()
        {
            return $"{StartColumn}{StartRow}:{EndColumn}{EndRow}";
        }
    }
}