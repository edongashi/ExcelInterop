using System;

namespace ExcelInterop
{
    public class ExitEventArgs : EventArgs
    {
        internal ExitEventArgs(ExitCause exitCause)
        {
            ExitCause = exitCause;
        }

        public ExitCause ExitCause { get; }
    }
}