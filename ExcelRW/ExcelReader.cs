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
    public delegate T CreateFromRow<T>(IRow row);
    /// <summary>
    /// ExcelReadr
    /// 使用NPOI读取Excel的辅助类
    /// </summary>
    public class ExcelReader
    {
        /// <summary>
        /// 将Isheet读入DataTable
        /// </summary>
        /// <param name="sheet">需要读取的sheet</param>
        /// <param name="skipBlankRow">是否跳过空行</param>
        /// <param name="headRowNum">标题行号，默认将第一行作为标题行</param>
        /// <param name="startRowNum">开始导入行号</param>
        /// <returns></returns>
        //TODO：添加是否设置DataTable列标题
        public static DataTable SheetToDataTable(ISheet sheet, bool skipBlankRow = false, int headRowNum = 0, int startRowNum = 0)
        {
            if (headRowNum >= 0 && headRowNum <= sheet.LastRowNum && startRowNum <= sheet.LastRowNum)
            {
                DataTable dt = new DataTable();
                int dtRowCount = sheet.LastRowNum + 1;
                IRow headRow = sheet.GetRow(headRowNum);
                int dtColCount = headRow.LastCellNum;
                //设置表头
                SetDTHead(dt, headRow, dtColCount);
                for (int i = startRowNum; i < dtRowCount; i++)
                {
                    DataRow dr = dt.NewRow();
                    IRow r = sheet.GetRow(i);
                    if (r != null)
                    {
                        for (int j = 0; j < dtColCount; j++)
                        {
                            dr[j] = CellValueToString(r.GetCell(j));
                        }
                    }
                    else
                    {
                        if (!skipBlankRow)
                        {
                            for (int j = 0; j < dtColCount; j++)
                            {
                                dr[j] = String.Empty;
                            }
                        }
                    }
                    dt.Rows.Add(dr);
                }
                return dt;
            }
            else
            {
                throw new ArgumentException("行数越界");
            }
        }
        /// <summary>
        /// 设置DataTable列标题
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="headRow"></param>
        /// <param name="dtColCount"></param>
        private static void SetDTHead(DataTable dt, IRow headRow, int dtColCount)
        {
            for (int i = 0; i < dtColCount; i++)
            {
                string colName = CellValueToString(headRow.GetCell(i));
                dt.Columns.Add(colName);
            }
        }
        /// <summary>
        /// 将Excel单元格转换为String
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string CellValueToString(ICell cell)
        {
            string cellString = String.Empty;

            if (cell != null)
            {
                switch (cell.CellType)
                {
                    case CellType.Numeric:
                        cellString = numericCellToString(cell);
                        break;
                    case CellType.String:
                        cellString = cell.StringCellValue;
                        break;
                    case CellType.Formula:
                        IFormulaEvaluator e = WorkbookFactory.CreateFormulaEvaluator(cell.Sheet.Workbook);
                        ICell formulaCell = e.EvaluateInCell(cell);
                        cellString = CellValueToString(formulaCell);
                        break;
                    case CellType.Boolean:
                        cellString = cell.BooleanCellValue.ToString();
                        break;
                    case CellType.Error:
                        cellString = cell.ErrorCellValue.ToString();
                        break;
                    case CellType.Blank:
                    case CellType.Unknown:
                    default:
                        break;
                }
            }
            return cellString;
        }
        /// <summary>
        /// 将数值类型单元格转换为String
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static string numericCellToString(ICell cell)
        {
            string numString = String.Empty;
            if (cell.CellType == CellType.Numeric)
            {
                if (IsDateCell(cell))
                {
                    numString = cell.DateCellValue.ToString();
                }
                else
                {
                    numString = cell.NumericCellValue.ToString();
                }
            }
            else
            {
                throw new ArgumentException("单元格格式错误");
            }
            return numString;
        }
        /// <summary>
        /// 判断单元格是否为日期格式
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        //部分自定义格式日期使用DateUtil.IsCellDateFormatted(Icell cell)判断时会出现错误
        public static bool IsDateCell(ICell cell)
        {
            bool isDate = false;

            string cellFormatString = cell.CellStyle.GetDataFormatString();

            if (DateUtil.IsCellDateFormatted(cell))
            {
                if (cellFormatString != null)
                {
                    isDate = !cellFormatString.Contains("General") && cellFormatString.Contains("DBNum");
                }
            }

            return isDate;
        }
        /// <summary>
        /// 将Sheet转换为List
        /// 在T中定义T CreateFromRow(IRow)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <param name="hasTitle"></param>
        /// <param name="CreateFromRow"></param>
        /// <returns></returns>
        public static List<T> SheetToList<T>(ISheet sheet,bool hasTitle,Func<IRow,T> CreateFromRow)
        { 
            List<T> list = new List<T>();
            int startNum = hasTitle ? 1 : 0;
            int count = sheet.GetRow(0).LastCellNum;
            for (int i = startNum; i < sheet.LastRowNum+1; i++)
            {
                IRow r = sheet.GetRow(i);
                T em = CreateFromRow(r);
                list.Add(em);
            }
            return list;
        }
        //使用定义的委托
        public static List<T> SheetToList<T>(ISheet sheet, CreateFromRow<T> Create)
        {
            List<T> list = new List<T>();
            int startNum = 0;
            for (int i = startNum; i < sheet.LastRowNum + 1; i++)
            {
                IRow r = sheet.GetRow(i);
                T em = Create(r);
                list.Add(em);
            }
            return list;
        }
        /// <summary>
        /// 将Row实例化为T
        /// </summary>
        /// <typeparam name="T">需要实例化的类</typeparam>
        /// <param name="row"></param>
        /// <returns></returns>
        public static T RowToModel<T>(IRow row)
        {
            if (row!=null)
            {
                Type model = typeof(T);
                T instance = (T)Activator.CreateInstance(model);
                foreach (var mprop in model.GetProperties())
                {
                    if (mprop.IsDefined(typeof(ColIndexAttribute)))
                    {
                        ColIndexAttribute ciAttr = mprop.GetCustomAttribute(typeof(ColIndexAttribute)) as ColIndexAttribute;
                        ICell cell = row.GetCell(ciAttr.Index);
                        mprop.SetValue(instance, ConvertCell(cell, mprop.PropertyType));
                        //mprop.SetValue(instance, row.GetCell(ciAttr.Index));
                    }
                    else
                    {
                        throw new Exception(string.Format("类型{0}属性{1}未定义ColIndex", model.Name, mprop.Name));
                    }
                }
                return instance;
            }
            else
            {
                return null;
            }
           
        }
        /// <summary>
        /// 将单元格的值转换为类型t
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="t"></param>
        /// <returns></returns>
        private static object ConvertCell(ICell cell,Type t)
        {
            switch (cell.CellType)
            {
                case CellType.Unknown:
                    return null;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return Convert.ChangeType(cell.DateCellValue, t);
                    }
                    else
                    {
                        return Convert.ChangeType(cell.NumericCellValue,t);
                    }
                case CellType.String:
                    return Convert.ChangeType(cell.StringCellValue, t);
                case CellType.Formula:
                    IFormulaEvaluator e = WorkbookFactory.CreateFormulaEvaluator(cell.Sheet.Workbook);
                    ICell formulaCell = e.EvaluateInCell(cell);
                    return ConvertCell(formulaCell, t);
                case CellType.Blank:
                    return null;
                case CellType.Boolean:
                    return Convert.ChangeType(cell.BooleanCellValue, t);
                case CellType.Error:
                    return null;
                default:
                    return null;
            }
        }

    }
}
