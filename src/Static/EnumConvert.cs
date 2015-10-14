using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInterop
{
    internal static class EnumConvert
    {
        public static Excel.XlInsertShiftDirection ConvertInsertShiftDirection(InsertShiftDirection shiftDirection)
        {
            return shiftDirection == InsertShiftDirection.ShiftDown
                ? Excel.XlInsertShiftDirection.xlShiftDown
                : Excel.XlInsertShiftDirection.xlShiftToRight;
        }

        public static Excel.XlDeleteShiftDirection ConvertDeleteShiftDirection(DeleteShiftDirection deleteDirection)
        {
            return deleteDirection == DeleteShiftDirection.ShiftUp
                ? Excel.XlDeleteShiftDirection.xlShiftUp
                : Excel.XlDeleteShiftDirection.xlShiftToLeft;
        }

        public static Excel.XlBorderWeight ConvertBorderThickness(BorderThickness thickness)
        {
            switch (thickness)
            {
                case BorderThickness.Thick:
                    return Excel.XlBorderWeight.xlThick;
                case BorderThickness.Medium:
                    return Excel.XlBorderWeight.xlMedium;
                default:
                    return Excel.XlBorderWeight.xlThin;
            }
        }
    }
}