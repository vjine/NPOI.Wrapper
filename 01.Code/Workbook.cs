using System;
using System.IO;
using NPOI.HPSF;
using NPOI.SS.UserModel;

//http://www.cnblogs.com/tonyqus/archive/2009/04/12/1434209.html
namespace NPOI.Wrapper
{
    [Serializable]
    public class Workbook
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

        public WorkSheet this[int index]
        {
            get
            {
                return this.Sheets[index];
            }
        }

        public WorkSheet this[string Name]
        {
            get
            {
                return this.Sheets[Name];
            }
        }

        private WorkSheetCollection _Sheets;
        public WorkSheetCollection Sheets
        {
            get
            {
                if (this._Sheets == null)
                {
                    this._Sheets = new WorkSheetCollection();
                    this._Sheets.bookHandler = this.bookHandler;

                    for (int i = 0; i < this.bookHandler.NumberOfSheets; i++)
                    {
                        ISheet sheet = this.bookHandler.GetSheetAt(i);
                        this.Sheets.Add(
                            new WorkSheet(sheet) { bookHandler = this.bookHandler, sheetHandler = sheet }
                            );
                    }
                }

                return this._Sheets;
            }
        }
    }
}
