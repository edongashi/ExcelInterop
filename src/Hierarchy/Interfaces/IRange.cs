using System;
using System.Drawing;

namespace ExcelInterop
{
    public interface IRange : ICellContainer
    {
        int EndColumn { get; }
        int EndRow { get; }
        int Height { get; }
        bool IsEmpty { get; }
        int StartColumn { get; }
        int StartRow { get; }
        int Width { get; }

        void CopyToLocation(int targetRow, int targetColumn);
        void CopyToLocation(IWorksheet targetWorksheet, int targetRow, int targetColumn);
        void Delete(DeleteShiftDirection deleteShiftDirection);
        void Expand(int rows, int columns);
        IRowRange GetEntireRows();
        IntPtr GetHemf();
        Image GetImage();
        IRange GetSubRange(int startRow, int startColumn, int height, int width);
        void InsertIntoLocation(int targetRow, int targetColumn, InsertShiftDirection shiftDirection);
        void InsertIntoLocation(IWorksheet targetWorksheet, int targetRow, int targetColumn, InsertShiftDirection shiftDirection);
        void Merge();
        void SetBackColor(Color color);
        void SetBorder(BorderCollection borders);
        void SetBorder(BorderThickness leftBorderThickness, BorderThickness topBorderThickness, BorderThickness rightBorderThickness, BorderThickness bottomBorderThickness);
        void SetFontColor(Color color);
        void SetTopBorderColor(Color color);
        void SetHorizontalBorderColor(Color color);
        void SetTopAndHorizontalBorderColor(Color color);
        void Shift(int rowOffset, int columnOffset);
    }
}