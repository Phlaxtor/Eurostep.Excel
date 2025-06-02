namespace Eurostep.Excel
{
    public sealed class ExcelSerializationSettings
    {
        public string SheetName { get; set; } = "Sheet1";

        public bool UseHeaders { get; set; } = true;
    }
}