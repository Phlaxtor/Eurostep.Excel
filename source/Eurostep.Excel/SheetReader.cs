using System.Diagnostics.CodeAnalysis;

namespace Eurostep.Excel
{
    internal sealed class SheetReader : ElementReader<ISheet, IRow>, ISheetReader
    {
        private readonly IDictionary<string, int> _nameToIndex = new Dictionary<string, int>();

        public SheetReader(ISheet sheet, ExcelContext context) : base(sheet, context)
        {
        }

        public int GetIndex(string header)
        {
            if (_nameToIndex.TryGetValue(header, out int index)) return index;
            throw new ArgumentException($"Provided value does not exist '{header}'", nameof(header));
        }

        public bool Read([NotNullWhen(true)] out ITableRow? value)
        {
            if (Read(out IRow? row))
            {
                value = row.GetTableRow(this);
                return true;
            }
            value = default;
            return false;
        }

        public bool Read<T>(Func<ITableRow, T> CreateRow, [NotNullWhen(true)] out T? value)
        {
            if (Read(out ITableRow? row))
            {
                value = CreateRow(row);
                return value != null;
            }
            value = default;
            return false;
        }

        public IEnumerable<T> ReadAll<T>(Func<ITableRow, T> CreateRow)
        {
            while (Enumerator.MoveNext())
            {
                ITableRow row = Enumerator.Current.GetTableRow(this);
                if (row.IsEmpty) continue;
                T value = CreateRow(row);
                if (value == null) continue;
                yield return value;
            }
        }

        public IEnumerable<ITableRow> ReadAllTableRows()
        {
            while (Enumerator.MoveNext())
            {
                yield return Enumerator.Current.GetTableRow(this);
            }
        }

        public bool ReadToHeaders(params string[] headers)
        {
            var hashToName = GetHashToNameLookup(headers);
            while (Read(out IRow? row))
            {
                var values = row.GetValues();
                if (values.Length < headers.Length) continue;
                int found = 0;
                for (int i = 0; i < values.Length; i++)
                {
                    string hash = values[i].GetToUpperWithoutWhiteSpace();
                    if (hashToName.TryGetValue(hash, out string? h) == false) continue;
                    found++;
                    _nameToIndex[h] = i;
                }
                if (found == headers.Length) return true;
            }
            return false;
        }

        private IDictionary<string, string> GetHashToNameLookup(string[] headers)
        {
            var hashToName = new Dictionary<string, string>();
            foreach (string header in headers)
            {
                string hash = header.GetToUpperWithoutWhiteSpace();
                hashToName[hash] = header;
            }
            return hashToName;
        }
    }
}