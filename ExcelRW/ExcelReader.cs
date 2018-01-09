using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using System.Data;

namespace ExcelRW
{
    public class ExcelReader
    {
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

        private static void SetDTHead(DataTable dt, IRow headRow, int dtColCount)
        {
            for (int i = 0; i < dtColCount; i++)
            {
                string colName = CellValueToString(headRow.GetCell(i));
                dt.Columns.Add(colName);
            }
        }

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
