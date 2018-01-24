using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using System.Data;
using System.Reflection;

namespace ExcelRW
{
    //另一种委托写法
    delegate IRow ModelToRow<T>(T t);
    /// <summary>
    /// NPOI写入Excel辅助类
    /// </summary>
    class ExcelWriter
    {
        #region DataTableToSheet
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
        #endregion
        #region ListToSheet 使用委托
        /// <summary>
        /// 将List转换为ISheet
        /// 在T中定义 IRow GetTitle(),IRow CreateRow(T)
        /// </summary>
        /// <param name="list"></param>
        /// <param name="sheet"></param>
        /// <param name="head">是否导入标题</param>
        /// <param name="CreateRow"></param>
        /// <param name="GetTitle"></param>
        public static void ListToSheet<T>(List<T> list,ISheet sheet,bool head,Func<IRow> GetTitle,Func<T,IRow> CreateRow)
        {
            int startNum = 0;
            if (head)
            {
                IRow headRow = sheet.CreateRow(startNum);
                headRow = GetTitle();
                startNum += 1;
            }
            foreach (var item in list)
            {
                IRow r = sheet.CreateRow(startNum);
                r = CreateRow(item);
            }
        }
        /// <summary>
        /// ListToSheet,在指定行号开始将List导入Sheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="sheet"></param>
        /// <param name="CreatRow"></param>
        /// <param name="startRowNum"></param>
        private static void ListToSheet<T>(List<T> list,ISheet sheet,Func<T,IRow> CreatRow,int startRowNum)
        {
            foreach (var item in list)
            {
                IRow r = sheet.CreateRow(startRowNum);
                r = CreatRow(item);
                startRowNum += 1;
            }
        }
        /// <summary>
        /// ListToSheet,将List导入Sheet，从第一行开始，覆盖原始数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="sheet"></param>
        /// <param name="CreateRow"></param>
        public static void ListToSheet<T>(List<T> list,ISheet sheet,Func<T,IRow> CreateRow)
        {
            ListToSheet(list,sheet,CreateRow,0);
        }
        /// <summary>
        /// ListToSheet,将List导入Sheet，添加到原始数据之后
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="sheet"></param>
        /// <param name="CreateRow"></param>
        public static void ListAppendToSheet<T>(List<T> list,ISheet sheet,Func<T,IRow> CreateRow)
        {
            int startNum = sheet.LastRowNum + 1;
            ListToSheet(list, sheet, CreateRow, startNum);
        }
        /// <summary>
        /// ListToSheet,将List导入Sheet，从第一行开始，覆盖原始数据,第一行为标题
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="sheet"></param>
        /// <param name="GetTitle"></param>
        /// <param name="CreateRow"></param>
        public static void ListToSheetWithTitle<T>(List<T> list, ISheet sheet,Func<IRow> GetTitle, Func<T, IRow> CreateRow)
        {
            int startNum = 0;
            IRow headRow = sheet.CreateRow(startNum);
            headRow = GetTitle();
            startNum += 1;
            ListAppendToSheet(list, sheet, CreateRow);
        }
        //另一种委托写法,需要测试
        public static void ListToSheet<T>(List<T> list,ISheet sheet,ModelToRow<T> MtoRow)
        {
            int startNum = 0;
            foreach (var item in list)
            {
                IRow r = sheet.CreateRow(startNum);
                
                r = MtoRow(item);
                startNum += 1;
            }
        }
        #endregion
        #region ModelToRow
        public static void ModelToRow<T>(T model,IRow row)
        {
            Type t = typeof(T);
            foreach (var item in t.GetProperties())
            {
                if (item.IsDefined(typeof(ColIndexAttribute))&&item.IsDefined(typeof(ColTypeAttribute)))
                {
                    ColIndexAttribute ciAttr = item.GetCustomAttribute<ColIndexAttribute>();
                    ICell cell = row.CreateCell(ciAttr.Index);

                    ColTypeAttribute ctAttr = item.GetCustomAttribute<ColTypeAttribute>();
                    switch (ctAttr.ColType)
                    {
                        case ColType.T_STR:
                            cell.SetCellValue((string)item.GetValue(model));
                            break;
                        case ColType.T_BOOL:
                            cell.SetCellValue((bool)item.GetValue(model));
                            break;
                        case ColType.T_DATE:
                            cell.SetCellValue((DateTime)item.GetValue(model));
                            break;
                        case ColType.T_NUM:
                            cell.SetCellValue((double)item.GetValue(model));
                            break;
                        default:
                            break;
                    }
                }      
            }
        }
        #endregion
    }
}
