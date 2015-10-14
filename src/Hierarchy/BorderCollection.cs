namespace ExcelInterop
{
    public struct BorderCollection
    {
        public static readonly BorderCollection None = new BorderCollection(BorderThickness.None, BorderThickness.None, BorderThickness.None, BorderThickness.None);

        public static readonly BorderCollection Thin = new BorderCollection(BorderThickness.Thin, BorderThickness.Thin, BorderThickness.Thin, BorderThickness.Thin);

        public BorderThickness LeftBorderThickness { get; set; }

        public BorderThickness TopBorderThickness { get; set; }

        public BorderThickness RightBorderThickness { get; set; }

        public BorderThickness BottomBorderThickness { get; set; }

        public BorderCollection(BorderThickness leftBorderThickness, BorderThickness topBorderThickness,
            BorderThickness rightBorderThickness, BorderThickness bottomBorderThickness)
        {
            LeftBorderThickness = leftBorderThickness;
            TopBorderThickness = topBorderThickness;
            RightBorderThickness = rightBorderThickness;
            BottomBorderThickness = bottomBorderThickness;
        }
    }
}