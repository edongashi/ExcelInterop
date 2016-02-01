using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterop
{
    public class Worksheet : IDisposable, IWorksheet
    {
        // Self
        internal Excel.Worksheet _worksheet;
        internal Excel.Range _cells;
        //

        private ObjectDisposedCallback disposeCallback;

        // State
        private bool displayGridlines;
        private bool disposed;

        internal Worksheet(
            ExcelApplication excel,
            Workbook workbook,
            Excel.Worksheet _worksheet,
            ObjectDisposedCallback disposeCallback,
            bool displayGridLines)
        {
            this.ExcelApplication = excel;
            this.Workbook = workbook;
            this._worksheet = _worksheet;
            this._cells = _worksheet.Cells;
            this.disposeCallback = disposeCallback;
            this.Name = _worksheet.Name;
            DisplayGridlines = displayGridLines;
        }

        public ExcelApplication ExcelApplication { get; private set; }

        public Workbook Workbook { get; }

        IWorkbook IWorksheet.Workbook => Workbook;

        public string Name { get; private set; }

        public bool IsDisposed => disposed;

        public IWorksheet Clone()
        {
            return Workbook.CloneWorksheet(this);
        }

        public bool DisplayGridlines
        {
            get
            {
                AssertNotDisposed();
                return displayGridlines;
            }
            set
            {
                AssertNotDisposed();
                _worksheet.Activate();
                Excel.Window activeWindow = ExcelApplication._excel.ActiveWindow;
                activeWindow.DisplayGridlines = value;
                Marshal.ReleaseComObject(activeWindow);
                displayGridlines = value;
            }
        }

        public int UsedWidth
        {
            get
            {
                AssertNotDisposed();
                Excel.Range usedRange = _worksheet.UsedRange;
                Excel.Range usedColumns = usedRange.Columns;
                var width = usedColumns.Count;
                Marshal.ReleaseComObject(usedColumns);
                Marshal.ReleaseComObject(usedRange);
                return width;
            }
        }

        public ICell Cell(int row, int column)
        {
            AssertNotDisposed();
            if (row < 0 || column < 0)
            {
                throw new ArgumentOutOfRangeException();
            }

            return new Cell(row, column, this);
        }

        public object[,] GetValues()
        {
            AssertNotDisposed();
            throw new NotImplementedException();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public IRowRange GetRows(int startRow, int endRow)
        {
            AssertNotDisposed();
            if (startRow > endRow)
            {
                var temp = endRow;
                endRow = startRow;
                startRow = temp;
            }

            if (startRow < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(startRow));
            }

            return new RowRange(startRow, endRow, this);
        }

        public IRange GetRange(int startRow, int startColumn, int endRow, int endColumn)
        {
            AssertNotDisposed();
            if (startRow > endRow)
            {
                return null;
            }

            if (startColumn > endColumn)
            {
                return null;
            }

            if (startRow < 0 || startColumn < 0)
            {
                throw new ArgumentOutOfRangeException(startRow < 0 ? nameof(startRow) : nameof(startColumn));
            }

            return new Range(startRow, startColumn, endRow, endColumn, this);
        }

        ~Worksheet()
        {
            Dispose(false);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
            {
                return;
            }

            if (disposing)
            {
                if (disposeCallback != null)
                {
                    disposeCallback(this);
                    disposeCallback = null;
                }
            }

            Marshal.ReleaseComObject(_cells);
            _cells = null;
            Marshal.ReleaseComObject(_worksheet);
            _worksheet = null;

            disposed = true;
        }

        protected virtual void AssertNotDisposed()
        {
            if (disposed)
            {
                throw new ObjectDisposedException(GetType().FullName);
            }
        }

        public void Activate()
        {
            AssertNotDisposed();
            _worksheet.Activate();
            _worksheet.Select();
        }
    }
}