using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterop
{
    public class ExcelApplication : IDisposable
    {
        // Interop
        private Excel.Application _excel;
        private Excel.Workbooks _workbooks;
        private ObjectDisposedCallback disposeCallback;

        // State
        private bool disposed;
        private Process excelProcess;
        private bool started;

        // Data
        private List<Workbook> workbooks;

        public ReadOnlyCollection<Workbook> Workbooks => workbooks?.AsReadOnly();

        public bool IsDisposed => disposed;

        public bool HasStarted => started;

        public void Dispose()
        {
            Dispose(true);
            OnExit(ExitCause.Disposed);
            GC.SuppressFinalize(this);
        }

        ~ExcelApplication()
        {
            Dispose(false);
            OnExit(ExitCause.GarbageCollected);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
            {
                return;
            }

            disposed = true;
            if (disposing && started)
            {
                int count = workbooks.Count;
                for (int i = 0; i < count; i++)
                {
                    workbooks[0].Dispose();
                }

                workbooks = null;
                excelProcess.Dispose();
            }

            if (started)
            {
                Marshal.ReleaseComObject(_workbooks);
                _workbooks = null;
                Process process = null;
                try
                {
                    var hWnd = (IntPtr)_excel.Hwnd;
                    process = ProcessFunctions.GetProcessByHwnd(hWnd);
                    _excel.DisplayAlerts = true;
                    _excel.IgnoreRemoteRequests = false;
                    _excel.Quit();
                    Marshal.ReleaseComObject(_excel);
                    _excel = null;
                    process.WaitForExit(1000);
                }
                catch
                {
                }
                finally
                {
                    if (process != null && !process.HasExited)
                    {
                        ProcessFunctions.TryKillProcess(process);
                    }

                    process?.Dispose();
                }
            }
        }

        protected virtual void AssertNotDisposed()
        {
            if (disposed)
            {
                throw new ObjectDisposedException(GetType().FullName);
            }
        }

        protected virtual void AssertStarted()
        {
            AssertNotDisposed();
            if (!started)
            {
                throw new InvalidOperationException("ExcelApplication instance is not running.");
            }
        }

        public Workbook Open(string filePath)
        {
            AssertStarted();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("Invalid file path.", nameof(filePath));
            }

            var workbook = workbooks.FirstOrDefault(w => w.FilePath == filePath);
            if (workbook == null)
            {
                try
                {
                    workbook = new Workbook(filePath, _workbooks.Open(filePath), disposeCallback, _excel);
                    workbooks.Add(workbook);
                }
                catch (Exception e)
                {
                    throw new Exception($"An error occured while opening \"{filePath}\".", e);
                }
            }

            return workbook;
        }

        public void Start()
        {
            Start(false, false);
        }

        public void Start(bool displayAlerts, bool ignoreRemoteRequests)
        {
            AssertNotDisposed();
            if (started)
            {
                return;
            }

            try
            {
                _excel = new Excel.Application();
            }
            catch (Exception e)
            {
                throw new Exception("Failed to start Excel.", e);
            }

            var hWnd = (IntPtr)_excel.Hwnd;
            uint processId;
            ProcessFunctions.GetWindowThreadProcessId(hWnd, out processId);
            excelProcess = Process.GetProcessById((int)processId);
            excelProcess.EnableRaisingEvents = true;
            excelProcess.Exited += (s, e) =>
            {
                if (!disposed)
                {
                    Dispose(true);
                    OnExit(ExitCause.Unknown);
                }
            };

            _excel.DisplayAlerts = displayAlerts;
            _excel.IgnoreRemoteRequests = ignoreRemoteRequests;

            _workbooks = _excel.Workbooks;
            workbooks = new List<Workbook>();
            disposeCallback = sender => workbooks.Remove((Workbook)sender);
            started = true;
        }

        public IntPtr GetHwnd()
        {
            AssertStarted();
            return (IntPtr)_excel.Hwnd;
        }

        public int GetProcessId()
        {
            AssertStarted();
            return excelProcess.Id;
        }

        public event EventHandler<ExitEventArgs> Exited;

        private void OnExit(ExitCause exitCause)
        {
            Exited?.Invoke(this, new ExitEventArgs(exitCause));
        }
    }
}