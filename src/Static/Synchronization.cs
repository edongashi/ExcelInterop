namespace ExcelInterop
{
    internal static class Synchronization
    {
        public static readonly object ClipboardSyncRoot = new object();
    }
}