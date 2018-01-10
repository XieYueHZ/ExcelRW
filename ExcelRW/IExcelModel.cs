using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRW
{
    /// <summary>
    /// Model与Excel转换
    /// 取消接口，使用委托实现模型类与Excel转换
    /// </summary>
    public interface IExcelModel
    {
        //SortedList<int, string> ColTitle { get; set; }
        /// <summary>
        /// 依据IRow创建实例
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        //Error(使用委托)
        //IExcelModel CreateFromRow(IRow row);
        /// <summary>
        /// 将实例转换为IRow
        /// </summary>
        /// <returns></returns>
        IRow ToRow();
        /// <summary>
        /// 创建标题行
        /// </summary>
        /// <returns></returns>
        //Error
        IRow GetHeadRow();
    }
}
