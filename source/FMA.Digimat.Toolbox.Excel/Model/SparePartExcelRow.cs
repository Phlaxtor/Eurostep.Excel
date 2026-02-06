using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FMA.Digimat.Toolbox.Excel.Model
{
    public class SparePartExcelRow : CommonExcelRow
    {
        public override string SheetName => ExcelSheetName.SparePart;

        public override Dictionary<string, string> HeadingsWithColumnNames => new() {
            { ColumnHeading.NcageCage                  ,  "A"  },
            { ColumnHeading.PartNumber                 ,  "B"  },
            { ColumnHeading.CompleteItemName           ,  "C"  },
            { ColumnHeading.TechnicalDescription       ,  "D"  },
            { ColumnHeading.ModelIdentification        ,  "E"  },
            { ColumnHeading.NSNGC                      ,  "F"  },
            { ColumnHeading.NIIN                       ,  "G"  },
            { ColumnHeading.MainEquipment              ,  "H"  },
            { ColumnHeading.PartOfSystem               ,  "I"  },
            { ColumnHeading.BaseUnitOfMeasure          ,  "J"  },
            { ColumnHeading.Weight                     ,  "K"  },
            { ColumnHeading.Length                     ,  "L"  },
            { ColumnHeading.Width                      ,  "M"  },
            { ColumnHeading.Height                     ,  "N"  },
            { ColumnHeading.SuppliedWithSerialNumber   ,  "O"  },
            { ColumnHeading.UnitPrice                  ,  "P"  },
            { ColumnHeading.Currency                   ,  "Q"  },
            { ColumnHeading.EstimatedDeliveryTime      ,  "R"  },
            { ColumnHeading.EstimatedDeliveryTimeUnit  ,  "S"  },
            { ColumnHeading.BuMGtin                    ,  "T"  },
            { ColumnHeading.POuMUnitOfMeasure          ,  "U"  },
            { ColumnHeading.POuMSupplier               ,  "V"  },
            { ColumnHeading.POuMNumberOfBuM            ,  "W"  },
            { ColumnHeading.POuMGtin                   ,  "X"  },
            { ColumnHeading.ShelfLifeLimit             ,  "Y"  },
            { ColumnHeading.ShelfLifeLimitUnit         ,  "Z"  },
            { ColumnHeading.Repairable                 ,  "AA" },
            { ColumnHeading.PointOfContact             ,  "AB" },
            { ColumnHeading.URLwebsite                 ,  "AC" },
            { ColumnHeading.ProductDataSheet           ,  "AD" },
            { ColumnHeading.Isotop                     ,  "AE" },
            { ColumnHeading.Activity                   ,  "AF" },
            { ColumnHeading.LaserClassification        ,  "AG" },
            { ColumnHeading.LaserWavelength            ,  "AH" },
            { ColumnHeading.BeamDivergence             ,  "AI" },
            { ColumnHeading.Effect                     ,  "AJ" },
            { ColumnHeading.PulseLength                ,  "AK" },
            { ColumnHeading.Energy                     ,  "AL" },
            { ColumnHeading.Frequency                  ,  "AM" },
            { ColumnHeading.FrequencyUnit              ,  "AN" },
            { ColumnHeading.HazardousClass             ,  "AO" },
            { ColumnHeading.UNnumber                   ,  "AP" },
            { ColumnHeading.USOrigin                   ,  "AQ" },
            { ColumnHeading.RadioactiveMaterial        ,  "AR" },
        };
    }
}
