using System.Reflection;

namespace FMA.Digimat.Toolbox.Excel.Model
{
    internal sealed class ColumnInfo : IComparable, IComparable<ColumnInfo>
    {
        public string Column { get; init; }

        public uint ColumnNo { get; init; }

        public ExcelCellStyle? CellStyle { get; init; }

        public ExcelCellStyle? ColumnStyle { get; init; }

        public ExcelCellStyle? HeaderStyle { get; init; }

        public string Heading { get; init; }

        public uint MergeCellCount { get; set; }

        public ExcelCellStyle? MergeCellStyle { get; init; }

        public string? MergeCellTitle { get; init; }

        public bool Optional { get; init; }

        public PropertyInfo Property { get; init; }

        public double Width { get; init; }

        public int CompareTo(ColumnInfo? other)
        {
            if (other is null)
            {
                return 1;
            }

            return ColumnNo.CompareTo(other.ColumnNo);
        }

        public int CompareTo(object? obj)
        {
            if (obj is ColumnInfo other)
            {
                return CompareTo(other);
            }

            return 1;
        }

        public override string ToString()
        {
            return $"{Property.Name}, '{Heading}' ({Column}, {Width})";
        }
    }
}