using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace FMA.Digimat.Toolbox.Excel.Model
{
    public abstract class CommonExcelRow : ExcelRow, ICommonExcelRowData, IAmmunitionsExcelRowData, IChemicalsExcelRowData, IClothingExcelRowData, ISparePartExcelRowData, IOtherItemExcelRowData
    {
        public CommonExcelRow() { }
        public override uint FirstDataRow => 5;
        public override uint HeaderRow => 2;
        public uint GroupHeaderRow => 1;
        public uint Subheader1Row => 3;
        public uint Subheader2Row => 4;
        public uint LastDataRow => 1004;

        [ExcelColumn(ColumnHeading.NcageCage)]
        public string NcageCage { get; set; }

        [ExcelColumn(ColumnHeading.PartNumber)]
        public string PartNumber { get; set; }

        [ExcelColumn(ColumnHeading.CompleteItemName)]
        public string CompleteItemName { get; set; }

        [ExcelColumn(ColumnHeading.TechnicalDescription)]
        public string TechnicalDescription { get; set; }

        [ExcelColumn(ColumnHeading.ModelIdentification)]
        public string ModelIdentification { get; set; }

        [ExcelColumn(ColumnHeading.NSNGC)]
        public string NSNGC { get; set; }

        [ExcelColumn(ColumnHeading.NIIN)]
        public string NIIN { get; set; }

        [ExcelColumn(ColumnHeading.PartOfSystem)]
        public string PartOfSystem { get; set; }

        [ExcelColumn(ColumnHeading.BaseUnitOfMeasure)]
        public string BaseUnitOfMeasure { get; set; }

        [ExcelColumn(ColumnHeading.Weight)]
        public string Weight { get; set; }

        [ExcelColumn(ColumnHeading.Length)]
        public string Length { get; set; }

        [ExcelColumn(ColumnHeading.Width)]
        public string Width { get; set; }

        [ExcelColumn(ColumnHeading.Height)]
        public string Height { get; set; }

        [ExcelColumn(ColumnHeading.SuppliedWithSerialNumber)]
        public string SuppliedWithSerialNumber { get; set; }

        [ExcelColumn(ColumnHeading.UnitPrice)]
        public string UnitPrice { get; set; }

        [ExcelColumn(ColumnHeading.Currency)]
        public string Currency { get; set; }

        [ExcelColumn(ColumnHeading.EstimatedDeliveryTime)]
        public string EstimatedDeliveryTime { get; set; }

        [ExcelColumn(ColumnHeading.EstimatedDeliveryTimeUnit)]
        public string EstimatedDeliveryTimeUnit { get; set; }

        [ExcelColumn(ColumnHeading.BuMGtin)]
        public string BuMGtin { get; set; }

        [ExcelColumn(ColumnHeading.POuMUnitOfMeasure)]
        public string POuMUnitOfMeasure { get; set; }

        [ExcelColumn(ColumnHeading.POuMSupplier)]
        public string POuMSupplier { get; set; }

        [ExcelColumn(ColumnHeading.POuMNumberOfBuM)]
        public string POuMNumberOfBuM { get; set; }

        [ExcelColumn(ColumnHeading.POuMGtin)]
        public string POuMGtin { get; set; }

        [ExcelColumn(ColumnHeading.ShelfLifeLimit)]
        public string ShelfLifeLimit { get; set; }

        [ExcelColumn(ColumnHeading.ShelfLifeLimitUnit)]
        public string ShelfLifeLimitUnit { get; set; }

        [ExcelColumn(ColumnHeading.Repairable)]
        public string Repairable { get; set; }

        [ExcelColumn(ColumnHeading.PointOfContact)]
        public string PointOfContact { get; set; }

        [ExcelColumn(ColumnHeading.URLwebsite)]
        public string URLwebsite { get; set; }

        [ExcelColumn(ColumnHeading.ProductDataSheet)]
        public string ProductDataSheet { get; set; }

        [ExcelColumn(ColumnHeading.USOrigin)]
        public string USOrigin { get; set; }

        [ExcelColumn(ColumnHeading.RadioactiveMaterial)]
        public string RadioactiveMaterial { get; set; }

        [ExcelColumn(ColumnHeading.MainEquipment, false)]
        public string MainEquipment { get; set; } = "0";

        [ExcelColumn(ColumnHeading.Isotop, false)]
        public string Isotop { get; set; }

        [ExcelColumn(ColumnHeading.Activity, false)]
        public string Activity { get; set; }

        [ExcelColumn(ColumnHeading.LaserClassification, false)]
        public string LaserClassification { get; set; }

        [ExcelColumn(ColumnHeading.LaserWavelength, false)]
        public string LaserWavelength { get; set; }

        [ExcelColumn(ColumnHeading.BeamDivergence, false)]
        public string BeamDivergence { get; set; }

        [ExcelColumn(ColumnHeading.Effect, false)]
        public string Effect { get; set; }

        [ExcelColumn(ColumnHeading.PulseLength, false)]
        public string PulseLength { get; set; }

        [ExcelColumn(ColumnHeading.Energy, false)]
        public string Energy { get; set; }

        [ExcelColumn(ColumnHeading.Frequency, false)]
        public string Frequency { get; set; }

        [ExcelColumn(ColumnHeading.FrequencyUnit, false)]
        public string FrequencyUnit { get; set; }

        [ExcelColumn(ColumnHeading.HazardousClass, false)]
        public string HazardousClass { get; set; }

        [ExcelColumn(ColumnHeading.UNnumber, false)]
        public string UNnumber { get; set; }

        [ExcelColumn(ColumnHeading.NetExplosiveContent, false)]
        public string NetExplosiveContent { get; set; }

        [ExcelColumn(ColumnHeading.TotalExplosiveContent, false)]
        public string TotalExplosiveContent { get; set; }

        [ExcelColumn(ColumnHeading.CompatibilityGroup, false)]
        public string CompatibilityGroup { get; set; }

        [ExcelColumn(ColumnHeading.Size, false)]
        public string Size { get; set; }

        [ExcelColumn(ColumnHeading.Gender, false)]
        public string Gender { get; set; }

        public Dictionary<string, string> GetFriendlyNameLookup()
        {
            var lookup = new Dictionary<string, string>();
            Type objtype = this.GetType();
            foreach (PropertyInfo p in objtype.GetProperties())
            {
                var fieldAttribute = p.GetCustomAttributes(false).FirstOrDefault(z => z is ExcelColumnAttribute) as ExcelColumnAttribute;
                if (fieldAttribute != null && HeadingsWithColumnNames.ContainsKey((fieldAttribute.Heading)))
                {
                    var columnName = HeadingsWithColumnNames[fieldAttribute.Heading];
                    lookup[p.Name] = $"cell {columnName}{RowId} ({fieldAttribute.Heading})";
                }
            }
            return lookup;
        }
    }
}
