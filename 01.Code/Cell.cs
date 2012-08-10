using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;
using System.Collections;

namespace NPOI.Wrapper
{
    public class Cell
    {
        internal Cell()
        {
        }

        internal ICell cellHandler { get; set; }

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

        public Cell SetValue(object V)
        {
            Cell.SetValue(this.cellHandler, V, null);
            return this;
        }

        public Cell SetValue(object V, string format)
        {
            Cell.SetValue(this.cellHandler, V, format);
            return this;
        }

        public static Cell SetValue(Cell srcCell, Cell dstDell)
        {
            Cell.SetValue(srcCell.cellHandler, dstDell.cellHandler);
            return dstDell;
        }

        internal static void SetValue(ICell dstCell, object v, string format)
        {
            if (dstCell.CellType == CellType.Unknown || dstCell.CellType == CellType.BLANK || dstCell.CellType == CellType.STRING)
            {
                if (string.IsNullOrEmpty(format))
                {
                    dstCell.SetCellValue(v.ToString());
                }
                else
                {
                    dstCell.SetCellValue(string.Format("{0:" + format + "}", v));
                }
            }
            else if (dstCell.CellType == CellType.BOOLEAN)
            {
                bool bValue = false;
                if (!bool.TryParse(v.ToString(), out bValue))
                {
                    throw new Exception(
                        string.Format("Parse Bool Value Error:{0} On R[{1}],C[{2}]",
                        v, dstCell.Row, dstCell.ColumnIndex));
                }
                dstCell.SetCellValue(bValue);
            }
            else if (dstCell.CellType == CellType.NUMERIC)
            {
                double dbValue = 0;
                DateTime dtValue = DateTime.Now;

                if (double.TryParse(v.ToString(), out dbValue))
                {
                    dstCell.SetCellValue(dbValue);
                }
                else if (DateTime.TryParse(v.ToString(), out dtValue))
                {
                    dstCell.SetCellValue(dtValue);
                }
                else
                {
                    throw new Exception(
                        string.Format("Parse Numeric And Date Value Error:{0} On R[{1}],C[{2}]", 
                        v, dstCell.RowIndex, dstCell.ColumnIndex));
                }
            }
            else
            {
                throw new Exception(
                    string.Format("Type[{0}] Conflict To Set Value On R[{1}],C[{2}]",
                    dstCell.CellType, dstCell.RowIndex, dstCell.ColumnIndex)
                    );
            }
        }

        internal static ICell SetValue(ICell srcCell, ICell dstCell)
        {
            if (srcCell == null)
            {
                dstCell.Row.RemoveCell(dstCell);
                return null;
            }

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

            return dstCell;
        }
    }

    public class CellCollection : IEnumerator<Cell>, IEnumerable<Cell>
    {
        internal WorkSheet Sheet = null;

        public Cell this[int rowIndex, int colIndex]
        {
            get
            {
                ICell cell = WorkSheet.GetCell(this.Sheet.sheetHandler, rowIndex, colIndex, false);
                if (cell == null)
                {
                    return null;
                }

                return new Cell() { cellHandler = cell };
            }
        }

        public Cell this[int rowIndex, string colName]
        {
            get
            {
                return this[rowIndex, WorkSheet.GetColIndexByName(colName)];
            }
        }

        #region Enumerator

        int rowIndex = -1, colIndex = -1;
        Cell _Current = null;
        public Cell Current
        {
            get
            {
                return this._Current;
            }
        }

        public void Dispose()
        {
            this.Reset();
        }

        object System.Collections.IEnumerator.Current
        {
            get { throw new NotImplementedException(); }
        }

        internal Range EnumerateRange { get; set; }
        IRow row = null;
        ICell cell = null;
        public bool MoveNext()
        {
            this.cell = null;
            if (this.EnumerateRange == null)
            {
                this.EnumerateRange = this.Sheet.UsedRange;
            }

            if (this.row != null)
            {
                goto Next_Cell;
            }
            if (this.rowIndex == -1)
            {
                this.rowIndex = this.EnumerateRange.Top - 1;
            }
        Next_Row: ;
            this.rowIndex += 1;
            if (this.rowIndex > this.EnumerateRange.Bottom)
            {
                return false;
            }
            this.row = this.Sheet.sheetHandler.GetRow(this.rowIndex);
            if (this.row == null)
            {
                goto Next_Row;
            }

            if (this.colIndex == -1)
            {
                this.colIndex = row.FirstCellNum - 1;
            }
        Next_Cell: ;
            this.colIndex += 1;
            if (this.colIndex > this.row.LastCellNum - 1)
            {
                this.row = null;
                this.colIndex = -1;
                goto Next_Row;
            }
            this.cell = row.GetCell(this.colIndex);
            if (this.cell == null)
            {
                goto Next_Cell;
            }

            this._Current = null;
            this._Current = new Cell() { cellHandler = this.cell };

            return true;
        }

        public void Reset()
        {
            this.EnumerateRange = null;
            this.row = null;
            this.cell = null;
            this.rowIndex = -1;
            this.colIndex = -1;
        }

        public IEnumerator<Cell> GetEnumerator()
        {
            return this as IEnumerator<Cell>;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this;
        }
        #endregion Enumerator
    }
}
