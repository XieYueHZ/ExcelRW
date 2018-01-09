using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRW
{
    interface IExcelModel
    {
        IExcelModel Create(IRow row);
        IRow ToRow();
        IRow GetHeadRow();
    }
}
