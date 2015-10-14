namespace ExcelInterop
{
    public interface IWorksheet : ICellContainer
    {
        bool DisplayGridlines { get; set; }
        bool IsDisposed { get; }
        string Name { get; }
        int UsedWidth { get; }

        void Activate();
        void Dispose();
        IRange GetRange(int startRow, int startColumn, int endRow, int endColumn);
        IRowRange GetRows(int startRow, int endRow);
    }
}