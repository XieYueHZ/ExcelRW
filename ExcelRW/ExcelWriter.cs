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
    public delegate IRow ModelToRow<T>(T t);
    /// <summary>
    /// NPOI写入Excel辅助类
    /// </summary>
    public class ExcelWriter
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
        #region ListToSheet
        /// <summary>
        /// ListToSheet,在指定行号开始将List导入Sheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">需要转换的List</param>
        /// <param name="sheet">要写入的ISheet</param>
        /// <param name="CreatRow">将对象转换为IRow</param>
        /// <param name="startRowNum">在第几行开始导入数据</param>
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
        /// <param name="list">需要转换的List</param>
        /// <param name="sheet">要写入的ISheet</param>
        /// <param name="CreateRow">将对象转换为IRow</param>
        public static void ListToSheet<T>(List<T> list,ISheet sheet,Func<T,IRow> CreateRow)
        {
            ListToSheet(list,sheet,CreateRow,0);
        }
        /// <summary>
        /// ListToSheet,将List导入Sheet，添加到原始数据之后
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">需要转换的List</param>
        /// <param name="sheet">要写入的ISheet</param>
        /// <param name="CreateRow">将对象转换为IRow</param>
        public static void ListAppendToSheet<T>(List<T> list,ISheet sheet,Func<T,IRow> CreateRow)
        {
            int startNum = sheet.LastRowNum + 1;
            ListToSheet(list, sheet, CreateRow, startNum);
        }
        /// <summary>
        /// ListToSheet,将List导入Sheet，从第一行开始，覆盖原始数据,第一行为标题
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">需要转换的List</param>
        /// <param name="sheet">要写入的ISheet</param>
        /// <param name="GetTitle">设置标题行</param>
        /// <param name="CreateRow">将对象转换为IRow</param>
        public static void ListToSheetWithTitle<T>(List<T> list, ISheet sheet,Func<IRow> GetTitle, Func<T, IRow> CreateRow)
        {
            int startNum = 0;
            IRow headRow = sheet.CreateRow(startNum);
            headRow = GetTitle();
            startNum += 1;
            ListAppendToSheet(list, sheet, CreateRow);
        }
        /// <summary>
        /// ListToSheet,另一种委托的写法，在指定行号开始将List导入Sheet
        /// </summary>
        /// 调整List<T>,ISheet参数顺序，避免出现冲突
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet">要写入的ISheet</param>
        /// <param name="list">需要转换的List</param>
        /// <param name="MtoRow">将对象转换为IRow</param>
        /// <param name="startRowNum">在第几行开始导入数据</param>
        public static void ListToSheet<T>(ISheet sheet, List<T> list,  ModelToRow<T> MtoRow,int startRowNum)
        {
            foreach (var item in list)
            {
                IRow r = sheet.CreateRow(startRowNum);
                r = MtoRow(item);
                startRowNum += 1;
            }
        }
        #endregion
        #region ListToSheet 使用反射将对象转化为IRow
        /// <summary>
        /// ListToSheet，在指定行号开始将List导入Sheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">需要转换的List</param>
        /// <param name="sheet">要写入的ISheet</param>
        /// <param name="startRowNum">在第几行开始导入数据</param>
        public static void ListToSheet<T>(List<T> list,ISheet sheet,int startRowNum)
        {
            foreach (var item in list)
            {
                IRow r = sheet.CreateRow(startRowNum);
                ObjectToRow(item, r);
                Console.WriteLine(r.GetCell(1).StringCellValue);
                startRowNum += 1;
            }

        }
        #endregion
        #region 辅助方法：使用自定义属性和反射获取标题行和数据行
        /// <summary>
        /// 使用反射将T类型对象转换为IRow
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj">需要转换为IRow的对象</param>
        /// <param name="row">转换完成后的IRow</param>
        public static void ObjectToRow<T>(T obj,IRow row)
        {
            Type t = typeof(T);
            foreach (var item in t.GetProperties())
            {
                if (item.IsDefined(typeof(ColIndexAttribute)))
                {
                    ColIndexAttribute ciAttr = item.GetCustomAttribute<ColIndexAttribute>();
                    ICell cell = row.CreateCell(ciAttr.Index);
                    if (item.IsDefined(typeof(ColTypeAttribute)))
                    {                     
                        ColTypeAttribute ctAttr = item.GetCustomAttribute<ColTypeAttribute>();
                        switch (ctAttr.ColType)
                        {
                            case ColType.T_STR:
                                cell.SetCellType(CellType.String);
                                cell.SetCellValue((string)item.GetValue(obj));
                                break;
                            case ColType.T_BOOL:
                                cell.SetCellType(CellType.Boolean);
                                cell.SetCellValue((bool)item.GetValue(obj));
                                break;
                            case ColType.T_DATE:
                                cell.SetCellType(CellType.Numeric);
                                cell.SetCellValue((DateTime)item.GetValue(obj));
                                break;
                            case ColType.T_NUM:
                                cell.SetCellType(CellType.Numeric);
                                cell.SetCellValue(Convert.ToDouble(item.GetValue(obj)));
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        cell.SetCellType(CellType.String);
                        cell.SetCellValue((item.GetValue(obj)).ToString());
                    }
                    
                }      
            }
        }
        /// <summary>
        /// 根据类型设定标题行
        /// </summary>
        /// <param name="t"></param>
        /// <param name="row"></param>
        public static void TypeToHeadRow(Type t,IRow row)
        {
            foreach (var item in t.GetProperties())
            {
                if (item.IsDefined(typeof(ColIndexAttribute)))
                {
                    ColIndexAttribute ciAttr = item.GetCustomAttribute<ColIndexAttribute>();
                    ICell cell = row.CreateCell(ciAttr.Index);
                    cell.SetCellType(CellType.String);
                    if (item.IsDefined(typeof(ColNameAttribute)))
                    {
                        ColNameAttribute cnAttr = item.GetCustomAttribute<ColNameAttribute>();
                        cell.SetCellValue(cnAttr.ColName);
                    }
                    else
                    {
                        cell.SetCellValue(item.Name);
                    }                   
                }
            }
        }
        #endregion
    }
}
