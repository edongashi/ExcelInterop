using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterop
{
    public class RowRange : IRowRange
    {
        internal RowRange(int startRow, int endRow, Worksheet parent)
        {
            StartRow = startRow;
            EndRow = endRow;
            Parent = parent;
        }

        public int EndRow { get; private set; }

        public int StartRow { get; private set; }

        public int Height => EndRow - StartRow + 1;

        internal Worksheet Parent { get; }

        public bool IsEmpty
        {
            get { throw new NotImplementedException(); }
            set { }
        }

        public ICell Cells(int row, int column)
        {
            AssertNotDisposed();
            if (row < 0 || column < 0)
            {
                throw new ArgumentOutOfRangeException(row < 0 ? nameof(row) : nameof(column));
            }

            row += StartRow;
            if (row > EndRow)
            {
                throw new ArgumentOutOfRangeException(nameof(row));
            }

            return new Cell(row, column, Parent);
        }

        public object[,] GetValues()
        {
            AssertNotDisposed();
            Excel.Range _range = _GetRange();
            object[,] values = _range.Value2;
            Marshal.ReleaseComObject(_range);
            return values;
        }

        protected virtual void AssertNotDisposed()
        {
            if (Parent.IsDisposed)
            {
                throw new ObjectDisposedException(nameof(Parent), "Containing Worksheet has been disposed.");
            }
        }

        public void Shift(int offset)
        {
            if (StartRow + offset < 0)
            {
                throw new ArgumentException();
            }

            StartRow += offset;
            EndRow += offset;
        }

        public void Delete()
        {
            AssertNotDisposed();
            Excel.Range _range = _GetRange();
            _range.Delete();
            Marshal.ReleaseComObject(_range);
        }

        public void CopyToLocation(IWorksheet targetWorksheet, int targetRow)
        {
            AssertNotDisposed();
            var targetRange = targetWorksheet.GetRows(targetRow, targetRow + EndRow - StartRow) as RowRange;
            if (targetRange == null)
            {
                throw new InvalidOperationException("Implementation of this method depends on another Office Interop wrapper.");
            }

            Excel.Range _range = _GetRange();
            Excel.Range _targetRange = targetRange._GetRange();
            _range.Copy(_targetRange);
            Marshal.ReleaseComObject(_targetRange);
            Marshal.ReleaseComObject(_range);
        }

        public void CopyToLocation(int targetRow)
        {
            CopyToLocation(Parent, targetRow);
        }

        public void InsertIntoLocation(IWorksheet targetWorksheet, int targetRow)
        {
            AssertNotDisposed();
            var targetRange = targetWorksheet.GetRows(targetRow, targetRow + EndRow - StartRow) as RowRange;
            if (targetRange == null)
            {
                throw new InvalidOperationException("Implementation of this method depends on another Office Interop wrapper.");
            }

            Excel.Range _range = _GetRange();
            Excel.Range _targetRange = targetRange._GetRange();
            lock (Synchronization.ClipboardSyncRoot)
            {
                _targetRange.Insert(CopyOrigin: _range.Copy());
            }

            Marshal.ReleaseComObject(_targetRange);
            Marshal.ReleaseComObject(_range);
        }

        public void InsertIntoLocation(int targetRow)
        {
            InsertIntoLocation(Parent, targetRow);
        }

        public void CopyDimensionsToLocation(IWorksheet targetWorksheet, int targetRow, bool copyContent)
        {
            AssertNotDisposed();
            var targetRange = targetWorksheet.GetRows(targetRow, targetRow + EndRow - StartRow) as RowRange;
            if (targetRange == null)
            {
                throw new InvalidOperationException("Implementation of this method depends on another Office Interop wrapper.");
            }

            Excel.Range _range = _GetRange();
            Excel.Range _targetRange = targetRange._GetRange();
            lock (Synchronization.ClipboardSyncRoot)
            {
                _range.Copy();
                _targetRange.PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths);
                if (copyContent)
                {
                    _targetRange.PasteSpecial();
                }
            }

            Marshal.ReleaseComObject(_targetRange);
            Marshal.ReleaseComObject(_range);
        }

        public void CopyDimensionsToLocation(int targetRow, bool copyContent)
        {
            CopyDimensionsToLocation(Parent, targetRow, copyContent);
        }

        private Excel.Range _GetRange()
        {
            Excel.Range _cells = Parent._Cells;
            Excel.Range _startRange = _cells[StartRow + 1, 1];
            Excel.Range _endRange = _cells[EndRow + 1, 1];
            Excel.Range _range = _cells.Range[_startRange, _endRange];
            Excel.Range _rowRange = _range.EntireRow;
            Marshal.ReleaseComObject(_range);
            Marshal.ReleaseComObject(_endRange);
            Marshal.ReleaseComObject(_startRange);
            return _rowRange;
        }
    }
}