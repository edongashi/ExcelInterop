namespace ExcelInterop
{
    public interface ICellContainer
    {
        ICell Cells(int row, int column);

        object[,] GetValues();
    }
}