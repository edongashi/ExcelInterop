using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcelInterop
{
    internal static class ProcessFunctions
    {
        public static readonly IntPtr MainWindowHandle = Process.GetCurrentProcess().MainWindowHandle;

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static Process GetProcessByHwnd(IntPtr hWnd)
        {
            uint processId;
            GetWindowThreadProcessId(hWnd, out processId);
            if (processId == 0)
            {
                throw new ArgumentException("Process has not been found by the given main window handle.", "hWnd");
            }

            return Process.GetProcessById((int)processId);
        }

        public static void KillProcessByMainWindowHwnd(IntPtr hWnd)
        {
            GetProcessByHwnd(hWnd).Kill();
        }

        public static bool TryKillProcess(Process process)
        {
            try
            {
                process.Kill();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool TryKillProcessByHwnd(IntPtr hWnd)
        {
            uint processId;
            GetWindowThreadProcessId(hWnd, out processId);
            if (processId == 0)
            {
                return false;
            }

            try
            {
                Process.GetProcessById((int)processId).Kill();
            }
            catch (ArgumentException)
            {
                return false;
            }
            catch (Win32Exception)
            {
                return false;
            }
            catch (NotSupportedException)
            {
                return false;
            }
            catch (InvalidOperationException)
            {
                return false;
            }

            return true;
        }
    }
}