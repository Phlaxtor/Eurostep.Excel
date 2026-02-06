using System.Reflection;

namespace FMA.Digimat.Toolbox.Excel.Model
{
    public sealed class DeliveryReportExcelRow : ExcelRow
    {
        public override string SheetName => ExcelSheetName.DeliveryReport;

        public override uint FirstDataRow => 12;

        public override uint HeaderRow => 11;

        public override Dictionary<string, string> HeadingsWithColumnNames => new() {
            { DeliveryReportHeading.PartIdentification         ,  "A"  },
            { DeliveryReportHeading.Name                       ,  "B"  },
            { DeliveryReportHeading.Description                ,  "C"  },
            { DeliveryReportHeading.Model                      ,  "D"  },
            { DeliveryReportHeading.Category                   ,  "E"  },
            { DeliveryReportHeading.ApprovalStatus             ,  "F"  },
            { DeliveryReportHeading.ApprovedOn                 ,  "G"  },
            { DeliveryReportHeading.LastUpdate                 ,  "H"  },
            { DeliveryReportHeading.System                     ,  "I"  },
            { DeliveryReportHeading.IsSerialized               ,  "J"  },
            { DeliveryReportHeading.IsEndItem                  ,  "K"  },
            { DeliveryReportHeading.NatoIdentification         ,  "L"  },
            { DeliveryReportHeading.PointOfContact             ,  "M"  },
            { DeliveryReportHeading.Url                        ,  "N"  },
            { DeliveryReportHeading.ProductDataSheet           ,  "O"  },
            { DeliveryReportHeading.Gtin                       ,  "P"  },
            { DeliveryReportHeading.UnitOfMeasure              ,  "Q"  },
            { DeliveryReportHeading.Weight                     ,  "R"  },
            { DeliveryReportHeading.SizeTotal                  ,  "S"  },
            { DeliveryReportHeading.PriceAndCurrency           ,  "T"  },
            { DeliveryReportHeading.PurchasedLeadTime          ,  "U"  },
            { DeliveryReportHeading.RecommendedShelfLife       ,  "V"  },
            { DeliveryReportHeading.Gtin                       ,  "W"  },
            { DeliveryReportHeading.UnitOfMeasure              ,  "X"  },
            { DeliveryReportHeading.Quantity                   ,  "Y"  },
            { DeliveryReportHeading.IsRepairable               ,  "Z"  },
            { DeliveryReportHeading.CountryOfOrigin            ,  "AA" },
            { DeliveryReportHeading.IsRadioactive              ,  "AB" },
            { DeliveryReportHeading.HazardClass                ,  "AC" },
            { DeliveryReportHeading.UnNumber                   ,  "AD" },
            { DeliveryReportHeading.Isotop                     ,  "AE" },
            { DeliveryReportHeading.Activity                   ,  "AF" },
            { DeliveryReportHeading.LaserClassification        ,  "AG" },
            { DeliveryReportHeading.LaserWavelenght            ,  "AH" },
            { DeliveryReportHeading.BeamDivergence             ,  "AI" },
            { DeliveryReportHeading.PeakEffect                 ,  "AJ" },
            { DeliveryReportHeading.PulseLength                ,  "AK" },
            { DeliveryReportHeading.Energy                     ,  "AL" },
            { DeliveryReportHeading.Frequency                  ,  "AM" },
            { DeliveryReportHeading.Size                       ,  "AN" },
            { DeliveryReportHeading.Gender                     ,  "AO" },
            { DeliveryReportHeading.CompatibilityGroup         ,  "AP" },
            { DeliveryReportHeading.NetExplosiveContent        ,  "AQ" },
            { DeliveryReportHeading.TotalExplosiveContent      ,  "AR" },
        };

        public bool IsUpdated { get; set; }

        [ExcelColumn(DeliveryReportHeading.PartIdentification, "A", 35, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public string PartIdentification { get; set; }

        [ExcelColumn(DeliveryReportHeading.Name, "B", 20, false)]
        [ExcelMergeColumns<MaterialIdentificationStyle>(DeliveryReportHeading.MaterialIdentification, 3)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public string Name { get; set; }

        [ExcelColumn(DeliveryReportHeading.Description, "C", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public string Description { get; set; }

        [ExcelColumn(DeliveryReportHeading.Model, "D", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public string Model { get; set; }

        [ExcelColumn(DeliveryReportHeading.Category, "E", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public string Category { get; set; }

        [ExcelColumn(DeliveryReportHeading.ApprovalStatus, "F", 20, false)]
        [ExcelMergeColumns<ApprovalInformationStyle>(DeliveryReportHeading.ApprovalInformation, 2)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public string ApprovalStatus { get; set; }

        [ExcelColumn(DeliveryReportHeading.ApprovedOn, "G", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public string ApprovedOn { get; set; }

        [ExcelColumn(DeliveryReportHeading.LastUpdate, "H", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public string LastUpdate { get; set; }

        [ExcelColumn(DeliveryReportHeading.System, "I", 20, false)]
        [ExcelMergeColumns<AdditionalInformationStyle>(DeliveryReportHeading.AdditionalInformation, 2)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue System { get; set; }

        [ExcelColumn(DeliveryReportHeading.IsSerialized, "J", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue IsSerialized { get; set; }

        [ExcelColumn(DeliveryReportHeading.IsEndItem, "K", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue IsEndItem { get; set; }

        [ExcelColumn(DeliveryReportHeading.NatoIdentification, "L", 20, false)]
        [ExcelMergeColumns<NatoCodificationStyle>(DeliveryReportHeading.NatoCodification, 3)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue NatoIdentification { get; set; }

        [ExcelColumn(DeliveryReportHeading.PointOfContact, "M", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue PointOfContact { get; set; }

        [ExcelColumn(DeliveryReportHeading.Url, "N", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue Url { get; set; }

        [ExcelColumn(DeliveryReportHeading.ProductDataSheet, "O", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue ProductDataSheet { get; set; }

        [ExcelColumn(DeliveryReportHeading.Gtin, "P", 20, false)]
        [ExcelMergeColumns<DefaultProvisioningPackageStyle>(DeliveryReportHeading.DefaultProvisioningPackage, 6)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue GtinDefault { get; set; }

        [ExcelColumn(DeliveryReportHeading.UnitOfMeasure, "Q", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue UnitOfMeasureDefault { get; set; }

        [ExcelColumn(DeliveryReportHeading.Weight, "R", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue Weight { get; set; }

        [ExcelColumn(DeliveryReportHeading.SizeTotal, "S", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue SizeTotal { get; set; }

        [ExcelColumn(DeliveryReportHeading.PriceAndCurrency, "T", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue PriceAndCurrency { get; set; }

        [ExcelColumn(DeliveryReportHeading.PurchasedLeadTime, "U", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue PurchasedLeadTime { get; set; }

        [ExcelColumn(DeliveryReportHeading.RecommendedShelfLife, "V", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue RecommendedShelfLife { get; set; }

        [ExcelColumn(DeliveryReportHeading.Gtin, "W", 20, false)]
        [ExcelMergeColumns<AdditionalProvisioningPackageStyle>(DeliveryReportHeading.AdditionalProvisioningPackage, 2)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue GtinAdditional { get; set; }

        [ExcelColumn(DeliveryReportHeading.UnitOfMeasure, "X", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue UnitOfMeasureAdditional { get; set; }

        [ExcelColumn(DeliveryReportHeading.Quantity, "Y", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue Quantity { get; set; }

        [ExcelColumn(DeliveryReportHeading.IsRepairable, "Z", 20, false)]
        [ExcelMergeColumns<RepairaibilityDisposalStyle>(DeliveryReportHeading.RepairaibilityDisposal, 2)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue IsRepairable { get; set; }

        [ExcelColumn(DeliveryReportHeading.CountryOfOrigin, "AA", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue CountryOfOrigin { get; set; }

        [ExcelColumn(DeliveryReportHeading.IsRadioactive, "AB", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue IsRadioactive { get; set; }

        [ExcelColumn(DeliveryReportHeading.HazardClass, "AC", 20, false)]
        [ExcelMergeColumns<HazardAndDangerousGoodsStyle>(DeliveryReportHeading.HazardAndDangerousGoods, 1)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue HazardClass { get; set; }

        [ExcelColumn(DeliveryReportHeading.UnNumber, "AD", 20, false)]
        public DeliveryReportValue UnNumber { get; set; }

        [ExcelColumn(DeliveryReportHeading.Isotop, "AE", 20, false)]
        [ExcelMergeColumns<RadioactiveSourceItemStyle>(DeliveryReportHeading.RadioactiveSourceItem, 1)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue Isotop { get; set; }

        [ExcelColumn(DeliveryReportHeading.Activity, "AF", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue Activity { get; set; }

        [ExcelColumn(DeliveryReportHeading.LaserClassification, "AG", 20, false)]
        [ExcelMergeColumns<LaserSourceItemStyle>(DeliveryReportHeading.LaserSourceItem, 5)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue LaserClassification { get; set; }

        [ExcelColumn(DeliveryReportHeading.LaserWavelenght, "AH", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue LaserWavelenght { get; set; }

        [ExcelColumn(DeliveryReportHeading.BeamDivergence, "AI", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue BeamDivergence { get; set; }

        [ExcelColumn(DeliveryReportHeading.PeakEffect, "AJ", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue PeakEffect { get; set; }

        [ExcelColumn(DeliveryReportHeading.PulseLength, "AK", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue PulseLength { get; set; }

        [ExcelColumn(DeliveryReportHeading.Energy, "AL", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue Energy { get; set; }

        [ExcelColumn(DeliveryReportHeading.Frequency, "AM", 20, false)]
        [ExcelMergeColumns<RadioFrequencySourceItemStyle>(DeliveryReportHeading.RadioFrequencySourceItem, 0)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue Frequency { get; set; }

        [ExcelColumn(DeliveryReportHeading.Size, "AN", 20, false)]
        [ExcelMergeColumns<ClothingItemStyle>(DeliveryReportHeading.ClothingItem, 1)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue SizeClothing { get; set; }

        [ExcelColumn(DeliveryReportHeading.Gender, "AO", 20, false)]
        public DeliveryReportValue Gender { get; set; }

        [ExcelColumn(DeliveryReportHeading.CompatibilityGroup, "AP", 20, false)]
        [ExcelMergeColumns<AmmunitionAndExplosiveItemStyle>(DeliveryReportHeading.AmmunitionAndExplosiveItem, 2)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue CompatibilityGroup { get; set; }

        [ExcelColumn(DeliveryReportHeading.NetExplosiveContent, "AQ", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue NetExplosiveContent { get; set; }

        [ExcelColumn(DeliveryReportHeading.TotalExplosiveContent, "AR", 20, false)]
        [ExcelHeaderStyle<DefaultHeaderStyle>]
        public DeliveryReportValue TotalExplosiveContent { get; set; }

        public override uint? GetRuntimeCellStyle(IExcelStyleCache styleCache, PropertyInfo property, string column, uint row)
        {
            if (IsUpdated == false)
            {
                return null;
            }

            var columnNo = column.GetColumnIndexFromName();
            if (columnNo < 8)
            {
                return null;
            }

            if (columnNo == 8)
            {
                return styleCache.GetCellStyle<UpdatedLastUpdateStyle>();
            }

            var value = property.GetValue(this) as DeliveryReportValue;
            if (value is not null && value.IsUpdated)
            {
                return styleCache.GetCellStyle<UpdatedReportValueStyle>();
            }

            return null;
        }
    }
}