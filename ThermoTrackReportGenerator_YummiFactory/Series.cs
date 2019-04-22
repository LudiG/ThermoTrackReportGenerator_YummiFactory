using System.Collections.Generic;

namespace ThermoTrackReportGenerator_YummiFactory
{
    public class Series
    {
        public string Name { get; set; }

        public List<float> Values { get; set; }

        public uint TrendlinePeriod { get; set; }

        public Series(string name, List<float> values, uint trendlinePeriod)
        {
            Name = name;
            Values = values;
            TrendlinePeriod = trendlinePeriod;
        }
    }
}