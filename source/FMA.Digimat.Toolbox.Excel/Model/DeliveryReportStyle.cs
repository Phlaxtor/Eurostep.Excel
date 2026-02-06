using DocumentFormat.OpenXml.Spreadsheet;

namespace FMA.Digimat.Toolbox.Excel.Model
{
    public sealed class AdditionalInformationStyle : ExcelCellStyle
    {
        public AdditionalInformationStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.AdditionalInformation;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class AdditionalProvisioningPackageStyle : ExcelCellStyle
    {
        public AdditionalProvisioningPackageStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.AdditionalProvisioningPackage;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class AmmunitionAndExplosiveItemStyle : ExcelCellStyle
    {
        public AmmunitionAndExplosiveItemStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.AmmunitionAndExplosiveItem;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class ApprovalInformationStyle : ExcelCellStyle
    {
        public ApprovalInformationStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.ApprovalInformation;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class ClothingItemStyle : ExcelCellStyle
    {
        public ClothingItemStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.ClothingItem;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class DefaultHeaderStyle : ExcelCellStyle
    {
        public DefaultHeaderStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FontStyle = new ReportFontStyle();
        }

        public sealed class ReportFontStyle : ExcelFontStyle
        {
            public ReportFontStyle()
            {
                IsBold = true;
            }
        }
    }

    public sealed class DefaultProvisioningPackageStyle : ExcelCellStyle
    {
        public DefaultProvisioningPackageStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.DefaultProvisioningPackage;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class HazardAndDangerousGoodsStyle : ExcelCellStyle
    {
        public HazardAndDangerousGoodsStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.HazardAndDangerousGoods;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class LaserSourceItemStyle : ExcelCellStyle
    {
        public LaserSourceItemStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.LaserSourceItem;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class MaterialIdentificationStyle : ExcelCellStyle
    {
        public MaterialIdentificationStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.MaterialIdentification;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class NatoCodificationStyle : ExcelCellStyle
    {
        public NatoCodificationStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.NatoCodification;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class RadioactiveSourceItemStyle : ExcelCellStyle
    {
        public RadioactiveSourceItemStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.RadioactiveSourceItem;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class RadioFrequencySourceItemStyle : ExcelCellStyle
    {
        public RadioFrequencySourceItemStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.RadioFrequencySourceItem;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class RepairaibilityDisposalStyle : ExcelCellStyle
    {
        public RepairaibilityDisposalStyle()
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportHeadingColor.RepairaibilityDisposal;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class ReportTitleStyle : ExcelCellStyle
    {
        public ReportTitleStyle()
        {
            Alignment = new Alignment { WrapText = true };
            FontStyle = new ReportFontStyle();
        }

        public sealed class ReportFontStyle : ExcelFontStyle
        {
            public ReportFontStyle()
            {
                FontSize = 18;
                IsBold = true;
            }
        }
    }

    public sealed class TitleHeaderFontStyle : ExcelFontStyle
    {
        public TitleHeaderFontStyle()
        {
            FontSize = 12;
            IsBold = true;
        }
    }

    public sealed class TitleHeaderTextFontStyle : ExcelFontStyle
    {
        public TitleHeaderTextFontStyle()
        {
            FontSize = 12;
        }
    }

    public sealed class UpdatedLastUpdateStyle : ExcelCellStyle
    {
        public UpdatedLastUpdateStyle()
        {
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportCellColor.UpdatedLastUpdate;
                PatternType = PatternValues.Solid;
            }
        }
    }

    public sealed class UpdatedReportValueStyle : ExcelCellStyle
    {
        public UpdatedReportValueStyle()
        {
            FillStyle = new ReportFillStyle();
        }

        public sealed class ReportFillStyle : ExcelFillStyle
        {
            public ReportFillStyle()
            {
                ForegroundColor = DeliveryReportCellColor.UpdatedReportValue;
                PatternType = PatternValues.Solid;
            }
        }
    }
}