namespace ThermoTrackReportGenerator_YummiFactory
{
    public class TableCell
    {
        public string Content { get; set; }

        public bool IsCentre { get; set; }
        public bool IsBold { get; set; }
        public int GridSpan { get; set; }

        public uint BorderTopSize { get; set; }
        public uint BorderLeftSize { get; set; }
        public uint BorderBottomSize { get; set; }
        public uint BorderRightSize { get; set; }

        public TableCell(string content, bool isCentre = false, bool isBold = false, int gridSpan = 1, uint borderTopSize = 0, uint borderLeftSize = 0, uint borderBottomSize = 0, uint borderRightSize = 0)
        {
            Content = content;

            IsCentre = isCentre;
            IsBold = isBold;
            GridSpan = gridSpan;

            BorderTopSize = borderTopSize;
            BorderLeftSize = borderLeftSize;
            BorderBottomSize = borderBottomSize;
            BorderRightSize = borderRightSize;
        }
    }
}
