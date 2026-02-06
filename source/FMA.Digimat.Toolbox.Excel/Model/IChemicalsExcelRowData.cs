namespace FMA.Digimat.Toolbox.Excel.Model
{
    public interface IChemicalsExcelRowData
    {
        public string NcageCage { get; set; }
        public string PartNumber { get; set; }
        public string CompleteItemName { get; set; }
        public string TechnicalDescription { get; set; }
        public string ModelIdentification { get; set; }
        public string NSNGC { get; set; }
        public string NIIN { get; set; }
        public string PartOfSystem { get; set; }
        public string BaseUnitOfMeasure { get; set; }
        public string Weight { get; set; }
        public string Length { get; set; }
        public string Width { get; set; }
        public string Height { get; set; }
        public string SuppliedWithSerialNumber { get; set; }
        public string UnitPrice { get; set; }
        public string Currency { get; set; }
        public string EstimatedDeliveryTime { get; set; }
        public string EstimatedDeliveryTimeUnit { get; set; }
        public string BuMGtin { get; set; }
        public string POuMUnitOfMeasure { get; set; }
        public string POuMSupplier { get; set; }
        public string POuMNumberOfBuM { get; set; }
        public string POuMGtin { get; set; }
        public string ShelfLifeLimit { get; set; }
        public string ShelfLifeLimitUnit { get; set; }
        public string Repairable { get; set; }
        public string PointOfContact { get; set; }
        public string URLwebsite { get; set; }
        public string ProductDataSheet { get; set; }
        public string USOrigin { get; set; }
        public string RadioactiveMaterial { get; set; }
        public string HazardousClass { get; set; }
        public string UNnumber { get; set; }
    }
}