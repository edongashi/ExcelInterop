using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterop
{
    public class Range : IRange
    {
        private const int CF_ETAFILE = 14;

        internal Range(int startRow, int startColumn, int endRow, int endColumn, Worksheet parent)
        {
            StartRow = startRow;
            EndRow = endRow;
            StartColumn = startColumn;
            EndColumn = endColumn;
            Parent = parent;
        }

        public int EndColumn { get; private set; }

        public int EndRow { get; private set; }

        public int StartColumn { get; private set; }

        public int StartRow { get; private set; }

        internal Worksheet Parent { get; }

        public int Width => EndColumn - StartColumn + 1;

        public int Height => EndRow - StartRow + 1;

        public bool IsEmpty
        {
            get
            {
                AssertNotDisposed();
                throw new NotImplementedException();
            }
        }

        public ICell Cell(int row, int column)
        {
            AssertNotDisposed();
            if (row < 0 || column < 0)
            {
                throw new ArgumentOutOfRangeException(row < 0 ? "row" : "column");
            }

            row += StartRow;
            column += StartColumn;
            if (row > EndRow || column > EndColumn)
            {
                throw new ArgumentOutOfRangeException(row > EndRow ? "row" : "column");
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

        public Image GetImage()
        {
            AssertNotDisposed();
            Excel.Range _range = _GetRange();
            Image image = null;
            lock (Synchronization.ClipboardSyncRoot)
            {
                _range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);
                if (Clipboard.ContainsImage())
                {
                    image = Clipboard.GetImage();
                }
            }

            Marshal.ReleaseComObject(_range);
            return image;
        }

        public IntPtr GetHemf()
        {
            AssertNotDisposed();
            Excel.Range _range = _GetRange();
            var hEmf = IntPtr.Zero;
            lock (Synchronization.ClipboardSyncRoot)
            {
                _range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);
                if (UnmanagedClipboard.OpenClipboard(ProcessFunctions.MainWindowHandle))
                {
                    if (UnmanagedClipboard.IsClipboardFormatAvailable(CF_ETAFILE) != 0)
                    {
                        hEmf = UnmanagedClipboard.GetClipboardData(CF_ETAFILE);
                    }

                    UnmanagedClipboard.CloseClipboard();
                }
            }

            Marshal.ReleaseComObject(_range);
            return hEmf;
        }

        public void SetBackColor(Color color)
        {
            AssertNotDisposed();
            Excel.Range _range = _GetRange();
            Excel.Interior _interior = _range.Interior;
            var oleColor = ColorTranslator.ToOle(color);
            _interior.Color = oleColor;
            Marshal.ReleaseComObject(_interior);
            Marshal.ReleaseComObject(_range);
        }

        public void SetFontColor(Color color)
        {
            AssertNotDisposed();
            Excel.Range _range = _GetRange();
            Excel.Font _font = _range.Font;
            var oleColor = ColorTranslator.ToOle(color);
            _font.Color = oleColor;
            Marshal.ReleaseComObject(_font);
            Marshal.ReleaseComObject(_range);
        }

        public void SetTopBorderColor(Color color)
        {
            Excel.Range _range = _GetRange();
            Excel.Borders _borders = _range.Borders;
            Excel.Border _border = _borders[Excel.XlBordersIndex.xlEdgeTop];
            var oleColor = ColorTranslator.ToOle(color);
            _border.Color = oleColor;
            Marshal.ReleaseComObject(_border);
            Marshal.ReleaseComObject(_borders);
            Marshal.ReleaseComObject(_range);
        }

        public void SetHorizontalBorderColor(Color color)
        {
            Excel.Range _range = _GetRange();
            Excel.Borders _borders = _range.Borders;
            Excel.Border _border = _borders[Excel.XlBordersIndex.xlInsideHorizontal];
            var oleColor = ColorTranslator.ToOle(color);
            _border.Color = oleColor;
            Marshal.ReleaseComObject(_border);
            Marshal.ReleaseComObject(_borders);
            Marshal.ReleaseComObject(_range);
        }

        public void SetTopAndHorizontalBorderColor(Color color)
        {
            Excel.Range _range = _GetRange();
            Excel.Borders _borders = _range.Borders;
            Excel.Border _borderTop = _borders[Excel.XlBordersIndex.xlEdgeTop];
            Excel.Border _borderHorizontal = _borders[Excel.XlBordersIndex.xlInsideHorizontal];
            var oleColor = ColorTranslator.ToOle(color);
            _borderTop.Color = oleColor;
            _borderHorizontal.Color = oleColor;
            Marshal.ReleaseComObject(_borderHorizontal);
            Marshal.ReleaseComObject(_borderTop);
            Marshal.ReleaseComObject(_borders);
            Marshal.ReleaseComObject(_range);
        }

        protected virtual void AssertNotDisposed()
        {
            if (Parent.IsDisposed)
            {
                throw new ObjectDisposedException(nameof(Parent), "Containing Worksheet has been disposed.");
            }
        }

        public void SetBorder(BorderCollection borders)
        {
            SetBorder(borders.LeftBorderThickness, borders.TopBorderThickness, borders.RightBorderThickness,
                borders.BottomBorderThickness);
        }

        public void SetBorder(BorderThickness leftBorderThickness, BorderThickness topBorderThickness,
            BorderThickness rightBorderThickness, BorderThickness bottomBorderThickness)
        {
            AssertNotDisposed();
            Excel.Range _range = _GetRange();
            var blackColor = ColorTranslator.ToOle(Color.Black);
            Excel.Borders _borders = _range.Borders;
            Excel.Border _leftBorder = _borders[Excel.XlBordersIndex.xlEdgeLeft];
            Excel.Border _topBorder = _borders[Excel.XlBordersIndex.xlEdgeTop];
            Excel.Border _rightBorder = _borders[Excel.XlBordersIndex.xlEdgeRight];
            Excel.Border _bottomBorder = _borders[Excel.XlBordersIndex.xlEdgeBottom];
            _leftBorder.LineStyle = leftBorderThickness == BorderThickness.None ? Excel.XlLineStyle.xlLineStyleNone : Excel.XlLineStyle.xlContinuous;
            _topBorder.LineStyle = topBorderThickness == BorderThickness.None ? Excel.XlLineStyle.xlLineStyleNone : Excel.XlLineStyle.xlContinuous;
            _rightBorder.LineStyle = rightBorderThickness == BorderThickness.None ? Excel.XlLineStyle.xlLineStyleNone : Excel.XlLineStyle.xlContinuous;
            _bottomBorder.LineStyle = bottomBorderThickness == BorderThickness.None ? Excel.XlLineStyle.xlLineStyleNone : Excel.XlLineStyle.xlContinuous;
            _leftBorder.Weight = EnumConvert.ConvertBorderThickness(leftBorderThickness);
            _leftBorder.Color = blackColor;
            _topBorder.Weight = EnumConvert.ConvertBorderThickness(topBorderThickness);
            _topBorder.Color = blackColor;
            _rightBorder.Weight = EnumConvert.ConvertBorderThickness(rightBorderThickness);
            _rightBorder.Color = blackColor;
            _bottomBorder.Weight = EnumConvert.ConvertBorderThickness(bottomBorderThickness);
            _bottomBorder.Color = blackColor;
            Marshal.ReleaseComObject(_bottomBorder);
            Marshal.ReleaseComObject(_rightBorder);
            Marshal.ReleaseComObject(_topBorder);
            Marshal.ReleaseComObject(_leftBorder);
            Marshal.ReleaseComObject(_borders);
            Marshal.ReleaseComObject(_range);
        }

        public void CopyToLocation(IWorksheet targetWorksheet, int targetRow, int targetColumn)
        {
            AssertNotDisposed();
            var targetRange = targetWorksheet.GetRange(targetRow, targetColumn, targetRow + EndRow - StartRow,
                targetColumn + EndColumn - StartColumn) as Range;
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

        public void CopyToLocation(int targetRow, int targetColumn)
        {
            CopyToLocation(Parent, targetRow, targetColumn);
        }

        public void InsertIntoLocation(IWorksheet targetWorksheet, int targetRow, int targetColumn,
            InsertShiftDirection shiftDirection)
        {
            AssertNotDisposed();
            var targetRange = targetWorksheet.GetRange(targetRow, targetColumn, targetRow + EndRow - StartRow,
                targetColumn + EndColumn - StartColumn) as Range;
            if (targetRange == null)
            {
                throw new InvalidOperationException("Implementation of this method depends on another Office Interop wrapper.");
            }

            Excel.Range _range = _GetRange();
            Excel.Range _targetRange = targetRange._GetRange();
            lock (Synchronization.ClipboardSyncRoot)
            {
                _targetRange.Insert(EnumConvert.ConvertInsertShiftDirection(shiftDirection), _range.Copy());
            }

            Marshal.ReleaseComObject(_targetRange);
            Marshal.ReleaseComObject(_range);
        }

        public void InsertIntoLocation(int targetRow, int targetColumn, InsertShiftDirection shiftDirection)
        {
            InsertIntoLocation(Parent, targetRow, targetColumn, shiftDirection);
        }

        public void Shift(int rowOffset, int columnOffset)
        {
            if (StartRow + rowOffset < 0 || StartColumn + columnOffset < 0)
            {
                throw new InvalidOperationException();
            }

            StartRow += rowOffset;
            EndRow += rowOffset;
            StartColumn += columnOffset;
            EndColumn += columnOffset;
        }

        public void Expand(int rows, int columns)
        {
            if (rows < 0 || columns < 0)
            {
                throw new ArgumentOutOfRangeException();
            }

            EndRow += rows;
            EndColumn += columns;
        }

        public IRange GetSubRange(int startRow, int startColumn, int height, int width)
        {
            if (startRow < 0 || startColumn < 0)
            {
                return null;
            }

            if (width == 0 || height == 0)
            {
                return null;
            }

            int subStartRow = this.StartRow + startRow;
            int subStartCol = this.StartColumn + startColumn;
            int subEndRow = subStartRow + height - 1;
            int subEndCol = subStartCol + width - 1;

            if (subStartRow > this.EndRow || subStartCol > this.EndColumn)
            {
                return null;
            }

            if (height < 0 || subEndRow > this.EndRow)
            {
                subEndRow = EndRow;
            }

            if (width < 0 || subEndCol > this.EndColumn)
            {
                subEndCol = EndColumn;
            }

            return new Range(subStartRow, subStartCol, subEndRow, subEndCol, Parent);
        }

        public IRowRange GetEntireRows()
        {
            AssertNotDisposed();
            return Parent.GetRows(StartRow, EndRow);
        }

        public void Delete(DeleteShiftDirection deleteShiftDirection)
        {
            AssertNotDisposed();
            Excel.Range _range = _GetRange();
            _range.Delete(EnumConvert.ConvertDeleteShiftDirection(deleteShiftDirection));
            Marshal.ReleaseComObject(_range);
        }

        public void Merge()
        {
            AssertNotDisposed();
            Excel.Range _range = _GetRange();
            _range.Merge();
            Marshal.ReleaseComObject(_range);
        }

        private Excel.Range _GetRange()
        {
            Excel.Range _cells = Parent._cells;
            Excel.Range _startRange = _cells[StartRow + 1, StartColumn + 1];
            Excel.Range _endRange = _cells[EndRow + 1, EndColumn + 1];
            Excel.Range _range = _cells.Range[_startRange, _endRange];
            Marshal.ReleaseComObject(_endRange);
            Marshal.ReleaseComObject(_startRange);
            return _range;
        }
    }
}