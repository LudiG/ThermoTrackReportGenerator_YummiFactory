using System.Collections.Generic;

namespace ThermoTrackReportGenerator_YummiFactory
{
    public class Series_Point
    {
        public string Name { get; set; }

        public List<KeyValuePair<double, double>> Values { get; set; }

        public uint TrendlinePeriod { get; set; }

        public Series_Point(string name, List<KeyValuePair<double, double>> values, uint trendlinePeriod)
        {
            Name = name;
            Values = values;
            TrendlinePeriod = trendlinePeriod;
        }
    }
}