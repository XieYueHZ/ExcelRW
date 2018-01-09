using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using System.Data;

namespace ExcelRW
{
    /// <summary>
    /// NPOI写入Excel辅助类
    /// </summary>
    class ExcelWriter
    {
        /// <summary>
        /// 将DataTable转化为Isheet
        /// </summary>
        /// <param name="dt">需要转换的DataTable</param>
        /// <param name="sheet">要写入的ISheet</param>
        /// <param name="head">是否将DataTable列标题作为ISheet标题</param>
        /// <param name="isAppend">是否将DataTable附加在ISheet末尾。</param>
        public static void DataTableToSheet(DataTable dt,ISheet sheet,bool head,bool isAppend)
        {
            int startNum = 0;
            if (isAppend)
            {
                startNum = sheet.LastRowNum + 1;
            }
            else
            {
                if (sheet.LastRowNum>0)
                {
                    for (int i = 0; i < sheet.LastRowNum; i++)
                    {
                        sheet.RemoveRow(sheet.GetRow(i));
                    }
                }
            }
            DataTableToSheet(dt,sheet,head,startNum);
        }
        /// <summary>
        /// 将DataTable转化为Isheet
        /// </summary>
        /// <param name="dt">需要转换的DataTable</param>
        /// <param name="sheet">要写入的ISheet</param>
        /// <param name="head">是否将DataTable列标题作为ISheet标题</param>
        /// <param name="startNum">开始写入ISheet的行号</param>
        private static void DataTableToSheet(DataTable dt,ISheet sheet,bool head,int startNum)
        {
            int colNum = dt.Columns.Count;
            int rowNum = dt.Rows.Count;
            if (head)
            {
                IRow headRow = sheet.CreateRow(startNum);
                for (int i = 0; i < colNum; i++)
                {
                    headRow.CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
                }
                startNum = startNum+1;
            }
            else
            {
                for (int i = 0; i < rowNum; i++)
                {
                    IRow r = sheet.CreateRow(i + startNum);
                    for (int j = 0; j < colNum; j++)
                    {
                        r.CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                    }
                }
            }
        }
        /// <summary>
        /// 将List转换为ISheet，需要实现IExcelModel借口
        /// </summary>
        /// <param name="list"></param>
        /// <param name="sheet"></param>
        /// <param name="head">是否导入标题</param>
        public static void ListToSheet(List<IExcelModel> list,ISheet sheet,bool head)
        {
            int startNum = 0;
            if (head)
            {
                IRow headRow = sheet.CreateRow(startNum);
                headRow = list[0].GetHeadRow();
                startNum += 1;
            }
            foreach (var item in list)
            {
                IRow r = sheet.CreateRow(startNum);
                r = item.ToRow();
            }
        }
    }
}
