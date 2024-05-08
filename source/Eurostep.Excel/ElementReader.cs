using System.Diagnostics.CodeAnalysis;

namespace Eurostep.Excel
{
    internal abstract class ElementReader<TEntity, TElement> : IReader<TEntity, TElement>
            where TEntity : class, ICollectionElement<TElement>
            where TElement : class, IElement
    {
        private bool _disposed;

        protected ElementReader(TEntity current, ExcelContext context)
        {
            Context = context;
            Current = current;
            Enumerator = current.GetEnumerator();
            EndOfReader = false;
        }

        public TEntity Current { get; }
        public bool EndOfReader { get; private set; }
        protected ExcelContext Context { get; }
        protected IEnumerator<TElement> Enumerator { get; }

        public void Dispose()
        {
            if (_disposed) return;
            Enumerator.Dispose();
            if (Context.LeaveOpen == false) Context.Dispose();
            _disposed = true;
        }

        public bool Read([NotNullWhen(true)] out TElement? value)
        {
            value = default;
            if (EndOfReader) return false;
            if (Enumerator.MoveNext() == false)
            {
                EndOfReader = true;
                return false;
            }
            value = Enumerator.Current;
            return true;
        }

        public IEnumerable<TElement> ReadAllElements()
        {
            while (Enumerator.MoveNext())
            {
                yield return Enumerator.Current;
            }
        }

        public TElement? ReadElement()
        {
            if (Read(out TElement? value))
            {
                return value;
            }
            return default;
        }
    }
}