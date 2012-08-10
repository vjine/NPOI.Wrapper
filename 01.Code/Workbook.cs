using System;
using System.IO;
using NPOI.HPSF;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Collections;

//http://www.cnblogs.com/tonyqus/archive/2009/04/12/1434209.html
namespace NPOI.Wrapper
{
    [Serializable]
    public class Workbook : IEnumerator<WorkSheet>, IEnumerable<WorkSheet>
    {
        DocumentSummaryInformation dsi = 
            PropertySetFactory.CreateDocumentSummaryInformation();
        SummaryInformation si = 
            PropertySetFactory.CreateSummaryInformation();

        private Workbook()
        {
            this.dsi.Company = "www.vJine.net";
            this.si.ApplicationName = "NPOI.Wraper";
            this.si.Author = "Ivan@vJine.net";
        }

        private NPOI.HSSF.UserModel.HSSFWorkbook bookHandler = null;
        public string FileName { get; private set; }

        #region Open And Close

        public static Workbook Create()
        {
            return Workbook.Create("");
        }

        public static Workbook Create(string FileName)
        {
            if (!string.IsNullOrEmpty(FileName) && File.Exists(FileName))
            {
                throw new WrapperException("File Already Exist:[{0}]", FileName);
            }

            Workbook book = new Workbook();

            book.FileName = FileName;
            FileStream handlerStream = null;
            try
            {
                handlerStream = new FileStream(FileName, FileMode.CreateNew);
                book.bookHandler = new NPOI.HSSF.UserModel.HSSFWorkbook();
            }
            finally
            {
                if (handlerStream != null)
                {
                    handlerStream.Close();
                }
            }

            return book;
        }

        public static Workbook Open(string FileName)
        {
            if (!File.Exists(FileName))
            {
                throw new WrapperException("There's no such file[{0}]", FileName);
            }

            Workbook book = new Workbook();
            book.FileName = FileName;
            FileStream handlerStream = null;
            try
            {
                handlerStream = new FileStream(FileName, FileMode.Open);
                book.bookHandler = new NPOI.HSSF.UserModel.HSSFWorkbook(handlerStream);
            }
            finally
            {
                if (handlerStream != null)
                {
                    handlerStream.Close();
                }
            }

            return book;
        }

        public void Save()
        {
            this.Save(this.FileName);
        }

        public void Save(string FileName)
        {
            if (string.IsNullOrEmpty(FileName))
            {
                throw new WrapperException("File Name Should Be Given!");
            }

            if (File.Exists(FileName))
            {
                throw new WrapperException("File Already Exist[{0}]", FileName);
            }

            Stream hStream = null;
            try
            {
                hStream = new FileStream(FileName, FileMode.CreateNew);

                this.bookHandler.Write(hStream);
            }
            finally
            {
                if (hStream != null)
                {
                    hStream.Close(); hStream = null;
                }
            }
        }

        public void Save(Stream stream)
        {
            this.bookHandler.Write(stream);
        }

        public void Close()
        {
            this.bookHandler = null;
        }
        #endregion

        public WorkSheet this[int Index]
        {
            get
            {
                this.CheckIndex(Index);

                ISheet sheet = this.bookHandler.GetSheetAt(Index);
                return new WorkSheet(this.bookHandler, sheet);
            }
        }

        public WorkSheet this[string Name]
        {
            get
            {
                ISheet sheet = this.bookHandler.GetSheet(Name);
                if (sheet == null)
                {
                    throw new WrapperException("工作表【{0}】不存在", Name);
                }

                return new WorkSheet(this.bookHandler, sheet);
            }
        }

        public int Count
        {
            get
            {
                return this.bookHandler.NumberOfSheets;
            }
        }

        public WorkSheet Copy(string srcName)
        {
            if (!this.Exist(srcName))
            {
                throw new WrapperException("源工作表【{0}】不存在", srcName);
            }

            return this.Copy(this[srcName]);
        }

        public WorkSheet Copy(WorkSheet srcSheet)
        {
            ISheet dstSheet =
                this.bookHandler.CloneSheet(this.bookHandler.NumberOfSheets - 1);

            WorkSheet sheet = new WorkSheet(this.bookHandler, dstSheet);

            return sheet;
        }

        public WorkSheet Copy(string srcName, string dstName)
        {
            if (!this.Exist(srcName))
            {
                throw new WrapperException("源工作表【{0}】不存在", srcName);
            }
            if (this.Exist(dstName))
            {
                throw new WrapperException("目标工作表【{0}】已存在", dstName);
            }

            WorkSheet srcSheet = this[srcName];
            return this.Copy(srcSheet, dstName);
        }

        public WorkSheet Copy(WorkSheet srcSheet, string dstName)
        {
            if (this.Exist(dstName))
            {
                throw new WrapperException("目标工作表【{0}】已存在", dstName);
            }

            WorkSheet sheet = this.Copy(srcSheet); sheet.Name = dstName;

            return sheet;
        }

        public WorkSheet Add(int Index)
        {
            this.CheckIndex(Index);

            WorkSheet wshNew = this.Add(this.GetDefaultSheetName());
            wshNew.Index = Index;

            return wshNew;
        }

        public WorkSheet Add(string Name)
        {
            if (this.Exist(Name))
            {
                throw new WrapperException("工作表【{0}】已存在",Name);
            }

            ISheet sheet = this.bookHandler.CreateSheet(Name);
            WorkSheet wshNew = new WorkSheet(this.bookHandler, sheet);

            return wshNew;
        }

        public void Remove(int Index)
        {
            this.CheckIndex(Index);

            this.bookHandler.RemoveSheetAt(Index);
        }

        public void Remove(string Name)
        {
            if (!this.Exist(Name))
            {
                throw new WrapperException("工作表【{0}】不存在", Name);
            }
            this.Remove(this[Name].Index);
        }

        public void Remove(WorkSheet sheet)
        {
            this.Remove(sheet.Index);
        }

        public void Clear()
        {
            while (this.bookHandler.NumberOfSheets > 0)
            {
                this.Remove(0);
            }
        }

        public WorkSheet Move(WorkSheet sheet, int Index)
        {
            this.CheckIndex(Index);

            sheet.Index = Index;
            return sheet;
        }

        public WorkSheet Move(WorkSheet sheet, string Name)
        {
            if (!this.Exist(Name))
            {
                throw new WrapperException("工作表【{0}】不存在", Name);
            }

            sheet.Index = this[Name].Index;

            return sheet;
        }

        public WorkSheet Move(WorkSheet sheet, WorkSheet sheetAfter)
        {
            sheet.Index = sheetAfter.Index;

            return sheet;
        }

        public bool Exsit(int Index)
        {
            this.CheckIndex(Index);

            return this.bookHandler.GetSheetAt(Index) != null;
        }

        public bool Exist(string Name)
        {
            return this.bookHandler.GetSheet(Name) != null;
        }

        static int MAX_OF_SHEETS = 1000;
        string GetDefaultSheetName()
        {
            for (int i = 1; i < MAX_OF_SHEETS; i++)
            {
                string Name = string.Format("Sheet({0})", i);
                if (!this.Exist(Name))
                {
                    return Name;
                }
            }

            throw new WrapperException("无法获取默认工作表名称");
        }

        void CheckIndex(int Index)
        {
            if (Index < 0 || Index > this.bookHandler.NumberOfSheets - 1)
            {
                throw new WrapperException("索引不在范围内【[0,{0}]】", this.bookHandler.NumberOfSheets - 1);
            }
        }

        #region Enumerator

        int sheetIndex = -1;
        public WorkSheet Current
        {
            get
            {
                return this[this.sheetIndex];
            }
        }

        public void Dispose()
        {
            this.sheetIndex = -1;
        }

        object IEnumerator.Current
        {
            get 
            {
                return this[this.sheetIndex];
            }
        }

        public bool MoveNext()
        {
            if (this.sheetIndex < this.bookHandler.NumberOfSheets - 1)
            {
                this.sheetIndex += 1;
                return true;
            }
            else
            {
                return false;
            }
        }

        public void Reset()
        {
            this.sheetIndex = -1;
        }

        public IEnumerator<WorkSheet> GetEnumerator()
        {
            return this as IEnumerator<WorkSheet>;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this;
        }
        #endregion Enumerator
    }
}
