using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterop
{
    public class Workbook : IDisposable
    {
        // Interop
        private Excel.Workbook _workbook;

        // Data
        private ObjectDisposedCallback disposeCallback;

        // State
        private bool disposed;
        private List<Worksheet> worksheets;

        internal Workbook(string filePath, Excel.Workbook _workbook, ObjectDisposedCallback disposeCallback,
            Excel.Application _excel)
        {
            this.FilePath = filePath;
            this._workbook = _workbook;
            this.disposeCallback = disposeCallback;
            worksheets = new List<Worksheet>();
            Excel.Sheets sheets = _workbook.Worksheets;
            ObjectDisposedCallback worksheetDisposeCallback = sender => worksheets.Remove((Worksheet)sender);
            for (int i = 1; i <= sheets.Count; i++)
            {
                worksheets.Add(new Worksheet(sheets[i], this, worksheetDisposeCallback, _excel));
            }

            worksheets[0].Activate();
            Marshal.ReleaseComObject(sheets);
        }

        public ReadOnlyCollection<Worksheet> Worksheets
        {
            get
            {
                AssertNotDisposed();
                return worksheets.AsReadOnly();
            }
        }

        public bool IsDisposed => disposed;

        public string FilePath { get; }

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
                int count = worksheets.Count;
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
                _workbook.Close(false);
            }
            catch
            {
            }

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
    }
}