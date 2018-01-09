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
        public static DataTable SheetToDataTable(ISheet sheet, bool skipBlankRow=false, int headRowNum=0, int startRowNum=0)
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
        private static string CellValueToString(ICell cell)
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
                        IFormulaEvaluator e = WorkbookFactory.CreateFormulaEvaluator(WB);
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
            if (cell.CellType==CellType.Numeric)
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
    }
}
