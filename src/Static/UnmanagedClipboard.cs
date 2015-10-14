using System;
using System.Runtime.InteropServices;

namespace ExcelInterop
{
    internal static class UnmanagedClipboard
    {
        [DllImport("user32.dll", EntryPoint = "OpenClipboard", SetLastError = true, ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall)]
        public static extern bool OpenClipboard(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "EmptyClipboard", SetLastError = true, ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall)]
        public static extern bool EmptyClipboard();

        [DllImport("user32.dll", EntryPoint = "SetClipboardData", SetLastError = true, ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr SetClipboardData(int uFormat, IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "CloseClipboard", SetLastError = true, ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall)]
        public static extern bool CloseClipboard();

        [DllImport("user32.dll", EntryPoint = "GetClipboardData", SetLastError = true, ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr GetClipboardData(int uFormat);

        [DllImport("user32.dll", EntryPoint = "IsClipboardFormatAvailable", SetLastError = true, ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall)]
        public static extern short IsClipboardFormatAvailable(int uFormat);
    }
}