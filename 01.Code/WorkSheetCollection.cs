using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace NPOI.Wrapper
{
    public class WorkSheetCollection : List<WorkSheet>
    {
        internal IWorkbook bookHandler = null;

        public new WorkSheet this[int index]
        {
            get
            {
                if (index <= base.Count)
                {
                    return base[index - 1];
                }
                else
                {
                    return null;
                }
            }
        }

        public WorkSheet this[string Name]
        {
            get
            {
                Name = Name.ToUpper();
                for (int i = 0; i < this.Count; i++)
                {
                    if (Name == base[i].Name.ToUpper())
                    {
                        return base[i];
                    }
                }

                return null;
            }
        }

        public WorkSheet Add(string Name)
        {
            ISheet sheet = this.bookHandler.CreateSheet(Name);
            WorkSheet wshNew = new WorkSheet(sheet) { bookHandler = this.bookHandler, sheetHandler = sheet };
            this.Add(wshNew);

            return wshNew;
        }
    }
}
