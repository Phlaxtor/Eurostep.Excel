namespace Eurostep.Excel
{
    public interface IElement : IDisposable
    {
        int Index { get; }
        bool IsEmpty { get; }
    }
}