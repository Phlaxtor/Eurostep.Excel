namespace FMA.Digimat.Toolbox.Excel.Model
{
    public sealed class ReportHeader
    {
        public string CreatedBy { get; set; }
        public DateTime CreatedOn { get; set; }
        public string Project { get; set; }
        public string Section { get; set; }
        public string Supplier { get; set; }
        public string Workpackage { get; set; }
    }
}