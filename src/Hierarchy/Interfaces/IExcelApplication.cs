using System;
using System.Collections.ObjectModel;

namespace ExcelInterop
{
    public interface IExcelApplication
    {
        ReadOnlyCollection<IWorkbook> Workbooks { get; }
        IWorkbook NewWorkbook();
        bool IsDisposed { get; }
        bool HasStarted { get; }
        void Close();
        IWorkbook Open(string filePath);
        void Start();
        void Start(bool visible, bool displayAlerts, bool ignoreRemoteRequests);
        IntPtr GetHwnd();
        int GetProcessId();
        event EventHandler<ExitEventArgs> Exited;
    }
}