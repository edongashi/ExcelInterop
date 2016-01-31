using System.Collections.ObjectModel;

namespace ExcelInterop
{
    public interface IWorkbook
    {
        ReadOnlyCollection<IWorksheet> Worksheets { get; }
        bool IsDisposed { get; }
        string FilePath { get; }
        string Name { get; }
        IWorksheet NewWorksheet();
        IWorksheet CloneWorksheet(IWorksheet worksheet);
        void Save();
        void SaveAs(string path);
        void Close(bool saveChanges);
        IWorksheet GetWorksheet(int number);
        int WorksheetCount { get; }
    }
}