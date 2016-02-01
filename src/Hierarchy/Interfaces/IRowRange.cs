using System.Drawing;

namespace ExcelInterop
{
    public interface IRowRange : ICellContainer
    {
        int EndRow { get; }
        bool IsEmpty { get; set; }
        int StartRow { get; }
        int Height { get; }
        string FontName { get; set; }
        int FontSize { get; set; }
        double RowHeight { get; set; }
        void SetBackColor(Color color);
        void CopyToLocation(int targetRow);
        void CopyToLocation(IWorksheet targetWorksheet, int targetRow);
        void Delete();
        void InsertIntoLocation(int targetRow);
        void InsertIntoLocation(IWorksheet targetWorksheet, int targetRow);
        void CopyDimensionsToLocation(int targetRow, bool copyContent);
        void CopyDimensionsToLocation(IWorksheet targetWorksheet, int targetRow, bool copyContent);
        void Shift(int offset);
    }
}