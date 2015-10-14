using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterop
{
    public class Cell : ICell
    {
        private readonly int column;
        private readonly int row;

        internal Cell(int row, int column, Worksheet parent)
        {
            this.row = row;
            this.column = column;
            this.Parent = parent;
        }

        public string Text
        {
            get
            {
                AssertNotDisposed();
                Excel.Range _cell = Parent._Cells[row + 1, column + 1];
                string text = _cell.Text;
                Marshal.ReleaseComObject(_cell);
                return text;
            }
            set
            {
                AssertNotDisposed();
                Excel.Range _cell = Parent._Cells[row + 1, column + 1];
                _cell.Value2 = value;
                Marshal.ReleaseComObject(_cell);
            }
        }

        public bool IsMerged
        {
            get
            {
                AssertNotDisposed();
                Excel.Range _cell = Parent._Cells[row + 1, column + 1];
                bool isMerged = _cell.MergeCells;
                Marshal.ReleaseComObject(_cell);
                return isMerged;
            }
        }

        internal Worksheet Parent { get; }

        protected virtual void AssertNotDisposed()
        {
            if (Parent.IsDisposed)
            {
                throw new ObjectDisposedException(nameof(Parent), "Containing Worksheet has been disposed.");
            }
        }
    }
}