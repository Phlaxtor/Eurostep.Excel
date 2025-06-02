namespace Eurostep.Excel
{
    public sealed class HeaderBuilder
    {
        private readonly ExcelWriter _excel;
        private readonly string _sheetName;
        private readonly List<IPresentationColumn> _headers = [];

        internal HeaderBuilder(ExcelWriter excel, string sheetName)
        {
            _excel = excel;
            _sheetName = sheetName;
        }

        public void Build()
        {
            _excel.SetCurrentTab(_sheetName);
            _excel.AddHeaders(_headers.ToArray());
        }

        public HeaderBuilder New(string displayName, double width, CellStyleValue? styleIndex = default, CellStyleValue? columnStyle = default)
        {
            _headers.Add(new DefaultPresentationColumn(displayName, width, styleIndex, columnStyle));
            return this;
        }
    }
}