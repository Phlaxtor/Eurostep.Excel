using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FMA.Digimat.Toolbox.Excel.Model
{
    public class ClothingExcelRow : CommonExcelRow
    {
        public override string SheetName => ExcelSheetName.Clothing;

        public override Dictionary<string, string> HeadingsWithColumnNames => new() {
            { ColumnHeading.NcageCage                  , "A"  },
            { ColumnHeading.PartNumber                 , "B"  },
            { ColumnHeading.CompleteItemName           , "C"  },
            { ColumnHeading.TechnicalDescription       , "D"  },
            { ColumnHeading.ModelIdentification        , "E"  },
            { ColumnHeading.NSNGC                      , "F"  },
            { ColumnHeading.NIIN                       , "G"  },
            { ColumnHeading.PartOfSystem               , "H"  },
            { ColumnHeading.BaseUnitOfMeasure          , "I"  },
            { ColumnHeading.Weight                     , "J"  },
            { ColumnHeading.Length                     , "K"  },
            { ColumnHeading.Width                      , "L"  },
            { ColumnHeading.Height                     , "M"  },
            { ColumnHeading.SuppliedWithSerialNumber   , "N"  },
            { ColumnHeading.UnitPrice                  , "O"  },
            { ColumnHeading.Currency                   , "P"  },
            { ColumnHeading.EstimatedDeliveryTime      , "Q"  },
            { ColumnHeading.EstimatedDeliveryTimeUnit  , "R"  },
            { ColumnHeading.BuMGtin                    , "S"  },
            { ColumnHeading.POuMUnitOfMeasure          , "T"  },
            { ColumnHeading.POuMSupplier               , "U"  },
            { ColumnHeading.POuMNumberOfBuM            , "V"  },
            { ColumnHeading.POuMGtin                   , "W"  },
            { ColumnHeading.ShelfLifeLimit             , "X"  },
            { ColumnHeading.ShelfLifeLimitUnit         , "Y"  },
            { ColumnHeading.Repairable                 , "Z"  },
            { ColumnHeading.PointOfContact             , "AA" },
            { ColumnHeading.URLwebsite                 , "AB" },
            { ColumnHeading.ProductDataSheet           , "AC" },
            { ColumnHeading.Size                       , "AD" },
            { ColumnHeading.Gender                     , "AE" },
            { ColumnHeading.USOrigin                   , "AF" },
            { ColumnHeading.RadioactiveMaterial        , "AG" },
        };
    }
}
