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
    }

    public class RowCollection : List<Row>
    {

    }
}
