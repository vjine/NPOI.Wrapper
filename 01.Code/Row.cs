using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace NPOI.Wrapper
{
    public class Row
    {
        internal IRow rowHandler = null;
        public int index
        {
            get
            {
                return this.rowHandler.RowNum;
            }
        }

        public void Remove(Cell cell)
        {
            this.rowHandler.RemoveCell(cell.cellHandler);
        }
    }

    public class RowCollection : List<Row>
    {

    }
}
