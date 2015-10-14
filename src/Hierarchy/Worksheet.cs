using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterop
{
    public class Worksheet : IDisposable, IWorksheet
    {
        // Interop
        private Excel.Range _cells;
        private Excel.Application _excel;
        private Excel.Worksheet _worksheet;
        private bool displayGridlines;

        // Data
        private ObjectDisposedCallback disposeCallback;

        // State
        private bool disposed;

        internal Worksheet(Excel.Worksheet _worksheet, Workbook parent, ObjectDisposedCallback disposeCallback,
            Excel.Application _excel)
        {
            this._worksheet = _worksheet;
            this.Parent = parent;
            this._cells = _worksheet.Cells;
            this.disposeCallback = disposeCallback;
            this._excel = _excel;
            this.Name = _worksheet.Name;
            this.displayGridlines = true;
        }

        public Workbook Parent { get; private set; }

        public string Name { get; private set; }

        public bool IsDisposed => disposed;

        internal Excel.Range _Cells => _cells;

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
                if (value != displayGridlines)
                {
                    _worksheet.Activate();
                    Excel.Window activeWindow = _excel.ActiveWindow;
                    activeWindow.DisplayGridlines = value;
                    Marshal.ReleaseComObject(activeWindow);
                    displayGridlines = value;
                }
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

        public ICell Cells(int row, int column)
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
                var temp = endRow;
                endRow = startRow;
                startRow = temp;
            }

            if (startColumn > endColumn)
            {
                var temp = endColumn;
                endColumn = startColumn;
                startColumn = temp;
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