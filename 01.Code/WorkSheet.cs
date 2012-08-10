using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;

namespace NPOI.Wrapper
{
    public class WorkSheet
    {
        internal IWorkbook bookHandler = null;
        internal ISheet sheetHandler = null;
        internal WorkSheet(IWorkbook bookHandler, ISheet sheetHandler)
        {
            this.bookHandler = bookHandler;
            this.sheetHandler = sheetHandler;

            this.Cells = new CellCollection() { Sheet = this };
        }

        public string Name
        {
            get
            {
                return this.sheetHandler.SheetName;
            }
            set
            {
                this.bookHandler.SetSheetName(this.bookHandler.GetSheetIndex(this.sheetHandler), value);
            }
        }

        public Range UsedRange
        {
            get
            {
                Range _Used = new Range();
                _Used.Top = this.sheetHandler.FirstRowNum;
                _Used.Bottom = this.sheetHandler.LastRowNum;
                IRow rowTop = this.sheetHandler.GetRow(_Used.Top);
                IRow rowBottom = this.sheetHandler.GetRow(_Used.Bottom);
                _Used.Left =
                    rowTop.FirstCellNum < rowBottom.FirstCellNum ? rowTop.FirstCellNum : rowBottom.FirstCellNum;
                _Used.Right =
                    rowTop.LastCellNum > rowBottom.LastCellNum ? rowTop.LastCellNum : rowBottom.LastCellNum;

                _Used.Right -= 1;

                return _Used;
            }
        }

        public int Index
        {
            get
            {
                return this.bookHandler.GetSheetIndex(this.sheetHandler);
            }
            set
            {
                if (Index < 0 || Index > this.bookHandler.NumberOfSheets -1)
                {
                    throw new WrapperException("索引不在范围内【[0,{0}]】", this.bookHandler.NumberOfSheets - 1);
                }
                this.bookHandler.SetSheetOrder(this.sheetHandler.SheetName, value);
            }
        }

        public Cell this[int row, int column]
        {
            get
            {
                return this.Cells[row, column];
            }
        }

        public Cell this[int row, string column]
        {
            get
            {
                return this.Cells[row, column];
            }
        }

        public CellCollection Cells { get; private set; }

        public RowCollection Rows { get; private set; }

        public ColumnCollection Columns { get; private set; }

        public void SetValue(object v, int rowIndex, int colIndex)
        {
            this.SetValue(v, null, rowIndex, colIndex);
        }

        public void SetValue(object v, string format, int rowIndex, int colIndex)
        {
            this.SetValue(WorkSheet.GetCell(this.sheetHandler, rowIndex, colIndex, true), v, format);
        }

        internal void SetValue(ICell cell, object v, string format)
        {
            Cell.SetValue(cell, v, format);
        }

        internal ICell SetValue(ICell cell, int rowIndex, int colIndex)
        {
            ICell cell2Set = WorkSheet.GetCell(this.sheetHandler, rowIndex, colIndex, true);
            return Cell.SetValue(cell, cell2Set);
        }

        internal static IRow GetRow(ISheet sheetHandler, int rowIndex)
        {
            return GetRow(sheetHandler, rowIndex, false);
        }

        internal static IRow GetRow(ISheet sheetHandler, int rowIndex, bool Create)
        {
            IRow row = sheetHandler.GetRow(rowIndex);
            if (row == null)
            {
                if (Create)
                {
                    row = sheetHandler.CreateRow(rowIndex);
                }
                else
                {
                    return null;
                }
            }

            return row;
        }

        internal static ICell GetCell(ISheet sheetHandler, int rowIndex, int colIndex)
        {
            return GetCell(sheetHandler, rowIndex, colIndex, false);
        }

        internal static ICell GetCell(ISheet sheetHandler, int rowIndex, int colIndex, bool Create)
        {
            IRow row = WorkSheet.GetRow(sheetHandler, rowIndex, Create);

            ICell cell = row.GetCell(colIndex);
            if (cell == null)
            {
                if (Create)
                {
                    cell = row.CreateCell(colIndex);
                }
                else
                {
                    return null;
                }
            }

            return cell;
        }

        #region Template

        public void Set(object obj)
        {
            Type tObj = obj.GetType();

            int rowIndexMin = this.sheetHandler.FirstRowNum;
            int rowIndexMax = this.sheetHandler.LastRowNum;

            object vContext = null;
            List<PropertyInfo> vProperties = new List<PropertyInfo>();
            List<ICell> vCells = new List<ICell>();
            List<string> vFormats = new List<string>();
            //Contents Except Tag
            List<ICell> vContents = new List<ICell>();

            bool IsList = false;
            IEnumerator rows = this.sheetHandler.GetRowEnumerator();
            while (rows.MoveNext())
            {//Enumerate All Cells To Find The Template TAG
                IRow row = rows.Current as IRow;
                foreach (ICell cell in row.Cells)
                {
                    string vTemplate = cell.StringCellValue;
                    if (!(vTemplate.Length >= 4 && vTemplate.IndexOf('{') == 0 && vTemplate.LastIndexOf('}') == vTemplate.Length - 1))
                    {
                        vContents.Add(cell);
                        vTemplate = null;
                        continue;
                    }

                    vTemplate = vTemplate.Replace(" ", "");
                    string pName = vTemplate.Substring(1, vTemplate.Length - 2);
                    string vFormat = "";
                    int formatIndex = vTemplate.LastIndexOf(':');
                    if (formatIndex > 0)
                    {
                        vFormat = pName.Substring(formatIndex);
                        pName = pName.Substring(0, formatIndex - 1);
                    }

                    object Context = new object();
                    PropertyInfo p = PropertyParser.GetProperty(obj, pName, out IsList, out Context);
                    if (p == null)
                    {
                        throw new WrapperException("Fail To Parse Property:[{0}]", pName);
                    }
                    else if (IsList)
                    {
                        if (vContext == null)
                        {
                            vContext = Context;
                        }
                        else if (Context != vContext)
                        {
                            throw new Exception(string.Format("Context Error [{0}]@[R[{1}],C[{2}]"));
                        }

                        vCells.Add(cell); vProperties.Add(p); vFormats.Add(vFormat);
                    }
                    else
                    {
                        object v = p.GetValue(Context, null);
                        this.SetValue(cell, v, vFormat);
                    }
                }
            }

            if (vContext == null)
            {
                return;
            }

            int rowStepOffSet = 0;
            int rowIndexStart = 0;
            int rowIndexStop = 0;
            {//Calculate the rowSetpOffset to Fit dumplicate rows TAG.
                for (int i = 0; i < vCells.Count; i++)
                {
                    if (i == 0)
                    {
                        rowIndexStart = vCells[i].RowIndex;
                        rowIndexStop = vCells[i].RowIndex;
                    }

                    if (vCells[i].RowIndex < rowIndexStart)
                    {
                        rowIndexStart = vCells[i].RowIndex;
                    }
                    if (vCells[i].RowIndex > rowIndexStop)
                    {
                        rowIndexStop = vCells[i].RowIndex;
                    }
                }
                rowStepOffSet = rowIndexStop - rowIndexStart + 1;
            }

            for (int i = 0; i < vContents.Count; i++)
            {
                if (vContents[i].RowIndex < rowIndexStart || vContents[i].RowIndex > rowIndexStop)
                {
                    vContents.RemoveAt(i); i--;
                }
            }

            IList vList = vContext as IList;
            for (int i = 0; i < vList.Count; i++)
            {
                for (int c = 0; c < vContents.Count; c++)
                {
                    this.SetValue(
                        vContents[c],
                        vContents[c].RowIndex + i * rowStepOffSet, vContents[c].ColumnIndex
                        );
                }

                for (int c = 0; c < vCells.Count; c++)
                {
                    this.SetValue(
                        vProperties[c].GetValue(vList[i], null), vFormats[c],
                        vCells[c].RowIndex + i * rowStepOffSet, vCells[c].ColumnIndex);
                }
            }
        }

        public void Copy(Range srcRange, Range dstRange)
        {
            bool IsRowShift = Range.IsRowShift(srcRange, dstRange);
            bool IsColShift = Range.IsColShift(srcRange, dstRange);

            for (int rowIndex = srcRange.Top; rowIndex <= srcRange.Bottom; rowIndex++)
            {
                int dstRowIndex = rowIndex - srcRange.Top + dstRange.Top;
                IRow srcRow = WorkSheet.GetRow(this.sheetHandler, rowIndex);
                if (srcRow == null)
                {
                    continue;
                }
                IRow dstRow = WorkSheet.GetRow(this.sheetHandler, dstRowIndex, true);
                if (IsRowShift)
                {
                    dstRow.Height = srcRow.Height;
                }

                for (int colIndex = srcRange.Left; colIndex <= srcRange.Right; colIndex++)
                {
                    int dstColIndex = colIndex - srcRange.Left + dstRange.Left;
                    ICell srcCell = srcRow.GetCell(colIndex);
                    ICell dstCell = this.SetValue(srcCell, dstRowIndex, dstColIndex);

                    if (IsColShift)
                    {
                        this.sheetHandler.SetColumnWidth(dstColIndex, this.sheetHandler.GetColumnWidth(colIndex));
                    }
                }
            }
        }

        public void Copy(Range srcRange, Cell dstCell)
        {
            
        }

        public void Clear(Range R)
        {
            this.Cells.Reset();

            this.Cells.EnumerateRange = R;
            IEnumerator<Cell> cells2Clear = this.Cells.GetEnumerator();
            while (cells2Clear.MoveNext())
            {
                cells2Clear.Current.cellHandler.Row.RemoveCell(cells2Clear.Current.cellHandler);
            }

            this.Cells.Reset();
        }
        #endregion Template

        public static int GetColIndexByName(string colName)
        {
            colName = colName.ToUpper();
            if (!Regex.IsMatch(colName, "^[A-Z]$"))
            {
                throw new WrapperException("列名【{0}】无效", colName);
            }
            
            int colIndex = 0;
            int charIndex = 0;
            for (int i = 0; i < colName.Length; i++)
            {
                charIndex = colName.Length - i - 1;
                colIndex += (int)Math.Pow(26, i) * (colName[charIndex] - 'A' + 1);
            }

            return colIndex - 1;
        }
    }
}
