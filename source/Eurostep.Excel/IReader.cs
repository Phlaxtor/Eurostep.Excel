using System.Diagnostics.CodeAnalysis;

namespace Eurostep.Excel
{
    public interface IReader<TEntity, TElement> : IDisposable
        where TEntity : class, IElement
        where TElement : class, IElement
    {
        TEntity Current { get; }
        bool EndOfReader { get; }

        bool Read([NotNullWhen(true)] out TElement? value);

        IEnumerable<TElement> ReadAllElements();

        TElement? ReadElement();
    }
}