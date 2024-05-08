using System.Diagnostics.CodeAnalysis;

namespace Eurostep.Excel
{
    public interface ISheetReader : IReader<ISheet, IRow>, IHeaderLookup
    {
        bool Read([NotNullWhen(true)] out ITableRow? value);

        bool Read<T>(Func<ITableRow, T> CreateRow, [NotNullWhen(true)] out T? value);

        IEnumerable<T> ReadAll<T>(Func<ITableRow, T> CreateRow);

        IEnumerable<ITableRow> ReadAllTableRows();

        bool ReadToHeaders(params string[] headers);
    }
}