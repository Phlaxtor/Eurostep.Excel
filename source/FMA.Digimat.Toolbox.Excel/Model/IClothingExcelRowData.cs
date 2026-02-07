namespace FMA.Digimat.Toolbox.Excel.Model
{
    public interface IClothingExcelRowData
    {
        string NcageCage { get; set; }
        string PartNumber { get; set; }
        string CompleteItemName { get; set; }
        string TechnicalDescription { get; set; }
        string ModelIdentification { get; set; }
        string NSNGC { get; set; }
        string NIIN { get; set; }
        string PartOfSystem { get; set; }
        string BaseUnitOfMeasure { get; set; }
        string Weight { get; set; }
        string Length { get; set; }
        string Width { get; set; }
        string Height { get; set; }
        string SuppliedWithSerialNumber { get; set; }
        string UnitPrice { get; set; }
        string Currency { get; set; }
        string EstimatedDeliveryTime { get; set; }
        string EstimatedDeliveryTimeUnit { get; set; }
        string BuMGtin { get; set; }
        string POuMUnitOfMeasure { get; set; }
        string POuMSupplier { get; set; }
        string POuMNumberOfBuM { get; set; }
        string POuMGtin { get; set; }
        string ShelfLifeLimit { get; set; }
        string ShelfLifeLimitUnit { get; set; }
        string Repairable { get; set; }
        string PointOfContact { get; set; }
        string URLwebsite { get; set; }
        string ProductDataSheet { get; set; }
        string USOrigin { get; set; }
        string RadioactiveMaterial { get; set; }
        string Size { get; set; }
        string Gender { get; set; }
    }
}