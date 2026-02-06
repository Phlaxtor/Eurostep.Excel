namespace FMA.Digimat.Toolbox.Excel.Model
{
    public static class ExcelType
    {
        public const string SparePart = "SparePart";
        public const string Ammunitions = "Ammunitions";
        public const string Chemicals = "Chemicals";
        public const string Clothing = "Clothing";
        public const string OtherItem = "OtherItem";
    }

    public static class ExcelSheetName
    {
        public const string SparePart = "Spare parts & components";
        public const string Ammunitions = "Ammunition and explosives";
        public const string Chemicals = "Chemicals and POL-products";
        public const string Clothing = "Clothing and gear";
        public const string Other = "Other supply items";
        public const string DeliveryReport = "Delivery help report";
    }

    public static class RowNumber
    {
        public const uint HeaderRow = 2;
        public const uint FirstDataRow = 5;
    }

    public static class ColumnHeading
    {
        public const string NcageCage = "NCAGE/CAGE*";
        public const string PartNumber = "Part number*";
        public const string CompleteItemName = "Complete item name*";
        public const string TechnicalDescription = "Technical description";
        public const string ModelIdentification = "Model identification";
        public const string NSNGC = "NSNGC";
        public const string NIIN = "NIIN";
        public const string MainEquipment = "Main equipment";
        public const string PartOfSystem = "Part of system";
        public const string BaseUnitOfMeasure = "Base Unit of Measure*";
        public const string Weight = "Weight* (kg)";
        public const string Length = "Length* (mm)";
        public const string Width = "Width* (mm)";
        public const string Height = "Height* (mm)";
        public const string SuppliedWithSerialNumber = "Supplied with serial number";
        public const string UnitPrice = "Unit price*";
        public const string Currency = "Currency*";
        public const string EstimatedDeliveryTime = "Estimated delivery time*";
        public const string EstimatedDeliveryTimeUnit = "Estimated delivery time unit*";
        public const string BuMGtin = "Global trade item number (GTIN) for base-UoM";
        public const string POuMUnitOfMeasure = "Purchase Order Unit of Measure";
        public const string POuMSupplier = "POuM supplier NCAGE/CAGE";
        public const string POuMNumberOfBuM = "# BuM in POuM";
        public const string POuMGtin = "Global trade item number (GTIN) for purchase order unit";
        public const string ShelfLifeLimit = "Shelf life limit";
        public const string ShelfLifeLimitUnit = "Shelf life limit unit";
        public const string Repairable = "Repairable*";
        public const string PointOfContact = "Point of contact / contact information";
        public const string URLwebsite = "URL / website";
        public const string ProductDataSheet = "Product data sheet";
        public const string Isotop = "Isotop";
        public const string Activity = "Activity (Bq)";
        public const string LaserClassification = "Laser classification";
        public const string LaserWavelength = "Laser wavelength (nm)";
        public const string BeamDivergence = "Beam divergence (mrad)";
        public const string Effect = "Effect (peak) (mW)";
        public const string PulseLength = "Pulse length (ns)";
        public const string Energy = "Energy (mJ)";
        public const string Frequency = "Frequency";
        public const string FrequencyUnit = "Frequency unit";
        public const string HazardousClass = "Hazardous class and dangerous goods";
        public const string UNnumber = "UN-number";
        public const string USOrigin = "US Origin";
        public const string RadioactiveMaterial = "Radioactive material";
        public const string NetExplosiveContent = "Net explosive content";
        public const string TotalExplosiveContent = "Total explosive content";
        public const string CompatibilityGroup = "Compatibility group";
        public const string Size = "Size";
        public const string Gender = "Gender";
    }

    public static class DefaultValue
    {
        public const double FontSize = 11;
        public const string BorderColor = "#000000";
        public const string FillColor = "#FFFFFF";
        public const string FontColor = "#000000";
        public const string FontName = "Calibri";
    }

    public static class ReportType
    {
        public const string DeliveryReport = "DeliveryReport";
    }

    public static class DeliveryReportHeading
    {
        public const string Activity = "Activity (Bq)";
        public const string AdditionalInformation = "Additional information";
        public const string AdditionalProvisioningPackage = "Additional provisionning package";
        public const string AmmunitionAndExplosiveItem = "Ammunitions and explosives";
        public const string ApprovalInformation = "Approval information";
        public const string ApprovalStatus = "Approval status";
        public const string ApprovedOn = "Approved on";
        public const string BeamDivergence = "Beam divergence (mrad)";
        public const string Category = "Category";
        public const string ClothingItem = "Clothing item";
        public const string CompatibilityGroup = "Compatibility group";
        public const string CountryOfOrigin = "Country of origin";
        public const string DefaultProvisioningPackage = "Default provisioning package";
        public const string Description = "Description";
        public const string Energy = "Energy (mJ)";
        public const string Frequency = "Frequency";
        public const string Gender = "Gender";
        public const string Gtin = "GTIN";
        public const string HazardAndDangerousGoods = "Hazard and dangerous goods";
        public const string HazardClass = "Hazard class";
        public const string IsEndItem = "Is end item";
        public const string Isotop = "Isotop";
        public const string IsRadioactive = "Is radioactive";
        public const string IsRepairable = "Is repairable";
        public const string IsSerialized = "Is serialized";
        public const string LaserClassification = "Laser classification";
        public const string LaserSourceItem = "Laser source item";
        public const string LaserWavelenght = "Laser wavelenght (nm)";
        public const string LastUpdate = "Last update";
        public const string MaterialIdentification = "Material identification";
        public const string Model = "Model";
        public const string Name = "Name";
        public const string NatoCodification = "NATO codification";
        public const string NatoIdentification = "NATO identification";
        public const string NetExplosiveContent = "Net explosive content (kg)";
        public const string PartIdentification = "Part identification";
        public const string PeakEffect = "Peak effect (mW)";
        public const string PointOfContact = "Point of contact";
        public const string PriceAndCurrency = "Price & currency";
        public const string ProductDataSheet = "Product data sheet";
        public const string PulseLength = "Pulse length (ns)";
        public const string PurchasedLeadTime = "Purchased lead time";
        public const string Quantity = "Quantity";
        public const string RadioactiveSourceItem = "Radioactivity source item";
        public const string RadioFrequencySourceItem = "Radiofrequency radiation source";
        public const string RecommendedShelfLife = "Recommended shelf life";
        public const string RepairaibilityDisposal = "Repairibility / disposal";
        public const string Size = "Size";
        public const string SizeTotal = "Size (LxWxH mm)";
        public const string System = "System";
        public const string TotalExplosiveContent = "Total explosive content (kg)";
        public const string UnitOfMeasure = "Unit of Measure";
        public const string UnNumber = "UN-number";
        public const string Url = "URL/Website";
        public const string Weight = "Weight (kg)";
    }

    public static class DeliveryReportHeadingColor
    {
        public const string AdditionalInformation = "FFB5E6A2";
        public const string AdditionalProvisioningPackage = "FF94DCF8";
        public const string AmmunitionAndExplosiveItem = "FFF2CEEF";
        public const string ApprovalInformation = "FFFBE2D5";
        public const string ClothingItem = "FFFBE2D5";
        public const string DefaultProvisioningPackage = "FFCAEDFB";
        public const string HazardAndDangerousGoods = "FFFFC000";
        public const string LaserSourceItem = "FFE4C994";
        public const string MaterialIdentification = "FFDAF2D0";
        public const string NatoCodification = "FFF2CEEF";
        public const string RadioactiveSourceItem = "FFEFDFBF";
        public const string RadioFrequencySourceItem = "FFF7EEDD";
        public const string RepairaibilityDisposal = "FFE49EDD";
    }

    public static class DeliveryReportCellColor
    {
        public const string UpdatedLastUpdate = "FFFF0000";
        public const string UpdatedReportValue = "FFFFEB9C";
    }
}