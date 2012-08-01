﻿using System;
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

            bool IsList = false;
            for (int rowIndex = rowIndexMin; rowIndex <= rowIndexMax; rowIndex++)
            {
                IRow row = this.sheetHandler.GetRow(rowIndex);
                if (row == null)
                {
                    continue;
                }

                foreach (ICell cell in row.Cells)
                {
                    if (cell == null)
                    {
                        continue;
                    }

                    string vTemplate = cell.StringCellValue;
                    vTemplate.Replace(" ", "");
                    if (!(vTemplate.Length >= 4 && vTemplate.IndexOf('{') == 0 && vTemplate.LastIndexOf('}') == vTemplate.Length - 1))
                    {
                        vTemplate = null; continue;
                    }

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

                        vCells.Add(cell);
                        vProperties.Add(p);
                        vFormats.Add(vFormat);
                    }
                    else
                    {
                        object v = p.GetValue(Context, null);
                        this.SetValue(cell, v, vFormat);
                    }
                }
            }

            if (vContext != null)
            {
                int rowStepOffSet = 0;
                {
                    int rowIndexStart = 0;
                    int rowIndexStop = 0;
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

                IList vList = vContext as IList;
                for (int i = 0; i < vList.Count; i += rowStepOffSet)
                {
                    for (int c = 0; c < vCells.Count; c++)
                    {
                        this.SetValue(
                            vCells[c].RowIndex + i, vCells[c].ColumnIndex,
                            vProperties[c].GetValue(vList[i], null), vFormats[c]
                            );
                    }
                }
            }
        }

        public void SetValue(int rowIndex, int colIndex, object v)
        {
            this.SetValue(rowIndex, colIndex, v, null);
        }

        public void SetValue(int rowIndex, int colIndex, object v,string format)
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

            this.SetValue(cell, v, format);
        }

        public void SetValue(ICell cell, object v, string format)
        {
            Cell.SetValue(cell, v, format);
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
