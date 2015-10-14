namespace ExcelInterop
{
    public interface ICell
    {
        bool IsMerged { get; }
        string Text { get; set; }
    }
}