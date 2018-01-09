using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRW
{
    /// <summary>
    /// Model与Excel转换借口
    /// </summary>
    interface IExcelModel
    {
        /// <summary>
        /// 依据IRow创建实例
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        IExcelModel Create(IRow row);
        /// <summary>
        /// 将实例转换为IRow
        /// </summary>
        /// <returns></returns>
        IRow ToRow();
        /// <summary>
        /// 创建标题行
        /// </summary>
        /// <returns></returns>
        IRow GetHeadRow();
    }
}
