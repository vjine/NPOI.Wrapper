using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;
using System.Reflection;
using System.Collections;

namespace NPOI.Wrapper
{
    public class WorkSheet
    {
        internal IWorkbook bookHandler = null;
        internal ISheet sheetHandler = null;
        internal WorkSheet(ISheet sheetHandler)
        {
            this.Cells = new CellCollection() { sheetHandler = sheetHandler };
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

        public int Index
        {
            get
            {
                return this.bookHandler.GetSheetIndex(this.sheetHandler);
            }
            set
            {
                this.bookHandler.SetSheetOrder(this.sheetHandler.SheetName, value - 1);
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

        public void SetValue(object v, int rowIndex, int colIndex)
        {
            this.SetValue(v, null, rowIndex, colIndex);
        }

        public void SetValue(object v, string format, int rowIndex, int colIndex)
        {
            this.SetValue(this.GetCell(rowIndex, colIndex), v, format);
        }

        public void SetValue(ICell cell, object v, string format)
        {
            Cell.SetValue(cell, v, format);
        }

        public void SetValue(ICell cell, int rowIndex, int colIndex)
        {
            ICell cell2Set = this.GetCell(rowIndex, colIndex);
            Cell.SetValue(cell, cell2Set);
        }

        public ICell GetCell(int rowIndex, int colIndex)
        {
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

            return cell;
        }

        public static int GetColIndexByName(string colName)
        {
            colName = colName.ToUpper();
            int colIndex = 0;
            int charIndex = 0;
            for (int i = 0; i < colName.Length; i++)
            {
                charIndex = colName.Length - i - 1;
                colIndex += (int)Math.Pow(26, i) * (colName[charIndex] - 'A' + 1);
            }

            return colIndex;
        }
    }
}
