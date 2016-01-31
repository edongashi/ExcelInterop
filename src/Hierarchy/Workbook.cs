using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterop
{
    public class Workbook : IDisposable, IWorkbook
    {
        // Self
        internal Excel.Workbook _workbook;
        internal Excel.Sheets _sheets;
        // 

        private ObjectDisposedCallback disposeCallback;
        private readonly ObjectDisposedCallback worksheetDisposeCallback;

        // State
        private bool disposed;
        private List<Worksheet> worksheets;

        internal Workbook(
            ExcelApplication excelApplication,
            Excel.Workbook _workbook,
            string filePath,
            ObjectDisposedCallback disposeCallback)
        {
            ExcelApplication = excelApplication;
            this._workbook = _workbook;
            this.FilePath = filePath;
            worksheets = new List<Worksheet>();
            _sheets = _workbook.Worksheets;
            worksheetDisposeCallback = sender => worksheets.Remove((Worksheet)sender);
            this.disposeCallback = disposeCallback;
            for (var i = 1; i <= _sheets.Count; i++)
            {
                worksheets.Add(new Worksheet(ExcelApplication, this, _sheets[i], worksheetDisposeCallback, false));
            }

            worksheets[0].Activate();
        }

        public ReadOnlyCollection<IWorksheet> Worksheets
        {
            get
            {
                AssertNotDisposed();
                return worksheets.Cast<IWorksheet>().ToList().AsReadOnly();
            }
        }

        public ExcelApplication ExcelApplication { get; }

        public bool IsDisposed => disposed;

        public string FilePath { get; }

        public string Name
        {
            get
            {
                AssertNotDisposed();
                return _workbook.Name;
            }
        }
        
        public void Dispose()
        {
            Dispose(true, false);
            GC.SuppressFinalize(this);
        }

        ~Workbook()
        {
            Dispose(false, false);
        }

        protected virtual void Dispose(bool disposing, bool saveChanges)
        {
            if (disposed)
            {
                return;
            }

            if (disposing)
            {
                var count = worksheets.Count;
                for (int i = 0; i < count; i++)
                {
                    worksheets[0].Dispose();
                }

                worksheets = null;
                if (disposeCallback != null)
                {
                    disposeCallback(this);
                    disposeCallback = null;
                }
            }

            try
            {
                _workbook.Close(saveChanges);
            }
            catch
            {
            }

            Marshal.ReleaseComObject(_sheets);
            _sheets = null;
            Marshal.ReleaseComObject(_workbook);
            _workbook = null;

            disposed = true;
        }

        protected virtual void AssertNotDisposed()
        {
            if (disposed)
            {
                throw new ObjectDisposedException(GetType().FullName);
            }
        }

        public IWorksheet NewWorksheet()
        {
            AssertNotDisposed();
            var _worksheet = _sheets.Add(After: worksheets[worksheets.Count - 1]._worksheet);
            var worksheet = new Worksheet(ExcelApplication, this, _worksheet, worksheetDisposeCallback, false);
            worksheets.Add(worksheet);
            return worksheet;
        }

        public IWorksheet CloneWorksheet(IWorksheet worksheet)
        {
            var item = worksheet as Worksheet;
            if (item == null)
            {
                throw new InvalidOperationException("Implementation of this method depends on another Office Interop wrapper.");
            }

            var index = worksheets.IndexOf(item);
            if (index == -1)
            {
                throw new InvalidOperationException("Specified worksheet does not belong to this workbook.");
            }

            var _worksheet = item._worksheet;
            _worksheet.Copy(After: _worksheet);
            var clone = new Worksheet(ExcelApplication, this, _sheets[index + 2], worksheetDisposeCallback, item.DisplayGridlines);
            worksheets.Insert(index + 1, clone);
            return clone;
        }

        public void Save()
        {
            AssertNotDisposed();
            _workbook.Save();
        }

        public void SaveAs(string path)
        {
            AssertNotDisposed();
            _workbook.SaveAs(path);
        }

        public void Close(bool saveChanges)
        {
            Dispose(true, saveChanges);
        }

        public IWorksheet GetWorksheet(int number) => worksheets[number - 1];

        public int WorksheetCount => worksheets.Count;
    }
}