using System.Text.RegularExpressions;

namespace Eurostep.Excel
{
    internal abstract class ExcelElement : IElement
    {
        private bool _disposed;
        private int? _index;
        private bool? _isEmpty;

        protected ExcelElement(ExcelContext context)
        {
            Context = context;
        }

        public int Index => ReturnIndex();

        public bool IsEmpty => ReturnIsEmpty();

        protected ExcelContext Context { get; }

        public virtual void Dispose()
        {
            if (_disposed) return;
            if (Context.LeaveOpen == false) Context.Dispose();
            _disposed = true;
        }

        protected int GetColumnIndex(string column)
        {
            uint columnNo = GetColumnNo(column);
            int index = ((int)columnNo) - 1;
            return index;
        }

        protected string GetColumnName(uint index)
        {
            string result = string.Empty;
            int r = (int)index;
            while (r > 0)
            {
                int i = (r % 26);
                r = (r / 26);
                char c = (char)(i + 64);
                result = $"{c}{result}";
            }
            return result;
        }

        protected uint GetColumnNo(string column)
        {
            uint result = 0;
            int position = 0;
            for (int i = column.Length - 1; i >= 0; i--)
            {
                int index = char.ToUpper(column[i]) - 64;
                result += (uint)(Math.Pow(26, position) * index);
                position++;
            }
            return result;
        }

        protected abstract int GetIndex();

        protected abstract bool GetIsEmpty();

        protected (string Column, uint Row) ParseReference(string reference)
        {
            Regex regex = new Regex("(?<Column>[a-zA-Z]*)(?<Row>[0-9]*)");
            Match match = regex.Match(reference);
            return (match.Groups["Column"].Value, uint.Parse(match.Groups["Row"].Value));
        }

        private int ReturnIndex()
        {
            if (_index.HasValue) return _index.Value;
            _index = GetIndex();
            return _index.Value;
        }

        private bool ReturnIsEmpty()
        {
            if (_isEmpty.HasValue) return _isEmpty.Value;
            _isEmpty = GetIsEmpty();
            return _isEmpty.Value;
        }
    }
}