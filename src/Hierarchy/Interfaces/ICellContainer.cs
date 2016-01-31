namespace ExcelInterop
{
    public interface ICellContainer
    {
        ICell Cell(int row, int column);

        object[,] GetValues();
    }
}