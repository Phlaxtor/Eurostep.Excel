using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eurostep.Excel
{
    public abstract class ExcelStylesheetAttribute<T> : ExcelAttribute
        where T : ExcelStylesheetDefinition
    {
        protected ExcelStylesheetAttribute()
        {
        }
    }
}
