namespace Eurostep.Excel
{
    public abstract class ExcelSerializer
    {
        private readonly Type _type;

        protected ExcelSerializer(Type type)
        {
            _type = type;
        }

        public static IEnumerable<T> Deserialize<T>(Stream stream)
        {
            ExcelSerializer<T> serializer = new ExcelSerializer<T>();
            return serializer.Deserialize(stream);
        }

        public static void Serialize<T>(Stream stream, IEnumerable<T> collection)
        {
            ExcelSerializer<T> serializer = new ExcelSerializer<T>();
            serializer.Serialize(stream, collection);
        }
    }

    public sealed class ExcelSerializer<T> : ExcelSerializer
    {
        public ExcelSerializer() : base(typeof(T))
        {
            //System.Xml.Serialization.XmlSerializer
        }

        public IEnumerable<T> Deserialize(Stream stream)
        {
            return Enumerable.Empty<T>();
        }

        public void Serialize(Stream stream, IEnumerable<T> collection)
        {
        }
    }
}