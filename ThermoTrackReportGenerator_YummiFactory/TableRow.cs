using System.Collections.Generic;

namespace ThermoTrackReportGenerator_YummiFactory
{
    public class TableRow
    {
        public List<TableCell> Cells { get; set; }

        public TableRow()
        {
            Cells = new List<TableCell>();
        }

        public TableRow(List<TableCell> cells)
        {
            Cells = cells;
        }
    }
}