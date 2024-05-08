namespace Eurostep.Excel
{
    public interface ICollectionElement<TElement> : IElement, IEnumerable<TElement>
        where TElement : class, IElement
    {
        int Length { get; }
    }
}