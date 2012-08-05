using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace NPOI.Wrapper
{
    public class Cell
    {
        internal ICell cellHandler = null;

        public string Value
        {
            get
            {
                return this.cellHandler.StringCellValue;
            }
            set
            {
                this.cellHandler.SetCellValue(value.ToString());
            }
        }

        public int rowIndex
        {
            get
            {
                return this.cellHandler.RowIndex;
            }
        }

        public int colIndex
        {
            get
            {
                return this.cellHandler.ColumnIndex;
            }
        }

        public static void SetValue(ICell cell, object v, string format)
        {
            if (cell.CellType == CellType.Unknown || cell.CellType == CellType.BLANK || cell.CellType == CellType.STRING)
            {
                if (string.IsNullOrEmpty(format))
                {
                    cell.SetCellValue(v.ToString());
                }
                else
                {
                    cell.SetCellValue(string.Format("{0:" + format + "}", v));
                }
            }
            else if (cell.CellType == CellType.BOOLEAN)
            {
                bool bValue = false;
                if (!bool.TryParse(v.ToString(), out bValue))
                {
                    throw new Exception(
                        string.Format("Parse Bool Value Error:{0} On R[{1}],C[{2}]",
                        v, cell.Row, cell.ColumnIndex));
                }
                cell.SetCellValue(bValue);
            }
            else if (cell.CellType == CellType.NUMERIC)
            {
                double dbValue = 0;
                DateTime dtValue = DateTime.Now;

                if (double.TryParse(v.ToString(), out dbValue))
                {
                    cell.SetCellValue(dbValue);
                }
                else if (DateTime.TryParse(v.ToString(), out dtValue))
                {
                    cell.SetCellValue(dtValue);
                }
                else
                {
                    throw new Exception(
                        string.Format("Parse Numeric And Date Value Error:{0} On R[{1}],C[{2}]", 
                        v, cell.RowIndex, cell.ColumnIndex));
                }
            }
            else
            {
                throw new Exception(
                    string.Format("Type[{0}] Conflict To Set Value On R[{1}],C[{2}]",
                    cell.CellType, cell.RowIndex, cell.ColumnIndex)
                    );
            }
        }

        public static void SetValue(ICell srcCell, ICell dstCell)
        {
            dstCell.SetCellType(srcCell.CellType);
            dstCell.CellStyle = srcCell.CellStyle;
            
            switch (srcCell.CellType)
            {
                case CellType.ERROR:
                    dstCell.SetCellErrorValue(srcCell.ErrorCellValue);
                    break;
                case CellType.BLANK:
                    break;
                case CellType.FORMULA:
                    dstCell.SetCellFormula(srcCell.CellFormula);
                    break;
                case CellType.BOOLEAN:
                    dstCell.SetCellValue(srcCell.BooleanCellValue);
                    break;
                case CellType.NUMERIC:
                    dstCell.SetCellValue(srcCell.NumericCellValue);
                    break;
                case CellType.STRING:
                    dstCell.SetCellValue(srcCell.StringCellValue);
                    break;
                case CellType.Unknown:
                    break;
            }
        }
    }

    public class CellCollection : Dictionary<string, Cell>
    {
        internal ISheet sheetHandler = null;

        public Cell this[int rowIndex, int colIndex]
        {
            get
            {
                rowIndex--; colIndex--;

                string Key = string.Format("{0}-{1}", rowIndex, colIndex);
                if (this.ContainsKey(Key))
                {
                    return this[Key];
                }

                IRow row = this.sheetHandler.GetRow(rowIndex);
                if (row == null)
                {
                    row = this.sheetHandler.CreateRow(rowIndex);
                }
                ICell cell = row.GetCell(colIndex);
                if (cell == null)
                {
                    cell = row.CreateCell(colIndex);
                }

                Cell C = new Cell() { cellHandler = cell };
                this.Add(Key, C);
                return C;
            }
        }

        public Cell this[int rowIndex, string colName]
        {
            get
            {
                return this[rowIndex, WorkSheet.GetColIndexByName(colName)];
            }
        }
    }
}
