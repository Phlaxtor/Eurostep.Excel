namespace FMA.Digimat.Toolbox.Excel.Model
{
    [AttributeUsage(AttributeTargets.Property)]
    internal class ExcelColumnAttribute : Attribute
    {
        public ExcelColumnAttribute(string heading)
        {
            Heading = heading;
        }

        public ExcelColumnAttribute(string heading, bool mandatory = true)
        {
            Heading = heading;
            Optional = !mandatory;
        }

        public ExcelColumnAttribute(string heading, string column, double width, bool mandatory = true)
        {
            Column = column;
            Heading = heading;
            Optional = !mandatory;
            Width = width;
        }

        internal virtual string Column { get; private set; }

        internal virtual string Heading { get; private set; }

        internal virtual bool Optional { get; private set; }

        internal virtual double Width { get; private set; } = 20;
    }
}