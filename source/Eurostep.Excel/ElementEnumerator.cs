using DocumentFormat.OpenXml;
using System.Collections;

namespace Eurostep.Excel
{
    internal abstract class ElementEnumerator<T> : IEnumerator<T>
             where T : IElement
    {
        private T? _current;

        protected ElementEnumerator(IEnumerator<OpenXmlElement> enumerator, ExcelContext context)
        {
            Enumerator = enumerator;
            Context = context;
            Init();
        }

        public T Current => _current ?? throw new ArgumentNullException(nameof(Current));
        object IEnumerator.Current => Current;
        protected ExcelContext Context { get; }
        protected IEnumerator<OpenXmlElement> Enumerator { get; }
        protected int Position { get; private set; }

        public void Dispose()
        {
            Init();
            Enumerator.Dispose();
        }

        public bool MoveNext()
        {
            while (Enumerator.MoveNext())
            {
                if (GetCurrent(out T? current))
                {
                    _current = current;
                    return true;
                }
            }
            Init();
            return false;
        }

        public void Reset()
        {
            Init();
            Enumerator.Reset();
        }

        protected abstract bool GetCurrent(out T? current);

        private void Init()
        {
            _current = default(T);
            Position = -1;
        }
    }
}