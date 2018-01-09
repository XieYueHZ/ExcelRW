using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using System.Data;

namespace ExcelRW
{
    class ExcelWriter
    {
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
