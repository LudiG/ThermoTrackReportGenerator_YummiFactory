using System;
using System.Collections.Generic;
using System.IO;

using MySql.Data.MySqlClient;

namespace ThermoTrackReportGenerator_YummiFactory
{
    class MySQLController
    {
        private string _sqlConnectionString;

        private string _logPath;

        /// <summary>
        /// Default class constructor.
        /// </summary>

        public MySQLController()
        {
            string sqlServerIP = Properties.Settings.Default.SQLServerIP;

            _sqlConnectionString = "server=" + sqlServerIP + ";uid=LudiG1601;pwd=jankeandbird4LIFE!;database=test;";

            _logPath = Path.Combine(Directory.GetCurrentDirectory(), @"Logs\MySQL\");
        }

        /// <summary>
        /// Method to query the MySQL database for a list of temperatures, by timestamp.
        /// </summary>
        /// <param name="beaconID">The beacon ID.</param>
        /// <param name="start">The start time for the data period.</param>
        /// <param name="end">The end time for the data period.</param>
        /// <param name="timestamp">The timestamp (for logging purposes).</param>

        public List<KeyValuePair<DateTime, float>> GetTemperaturesByTime(string beaconID, DateTime start, DateTime end, DateTime timestamp)
        {
            List<KeyValuePair<DateTime, float>> points = new List<KeyValuePair<DateTime, float>>();

            try
            {
                using (MySqlConnection sqlConnection = new MySqlConnection(_sqlConnectionString))
                {
                    sqlConnection.Open();

                    string sqlCommandString = "SELECT UNIX_TIMESTAMP(tsLogged), fTemperature " +
                                              "FROM tblbledata_raw " +
                                              "WHERE (szBeaconID = \"" + beaconID + "\") " +
                                              "AND tsLogged BETWEEN \"" + start.ToString("yyyy-MM-dd HH:mm:ss") + "\" AND \"" + end.ToString("yyyy-MM-dd HH:mm:ss") + "\" " +
                                              "ORDER BY tsLogged";

                    using (MySqlCommand sqlCommand = new MySqlCommand(sqlCommandString, sqlConnection))
                    {
                        using (MySqlDataReader sqlReader = sqlCommand.ExecuteReader())
                        {
                            while (sqlReader.Read())
                            {
                                points.Add(new KeyValuePair<DateTime, float>(DateTimeOffset.FromUnixTimeSeconds(sqlReader.GetUInt32(0)).UtcDateTime, sqlReader.GetFloat(1)));
                            }
                        }
                    }

                    sqlConnection.Close();
                }
            }

            catch (MySqlException exception)
            {
                Directory.CreateDirectory(_logPath);

                using (StreamWriter logWriter = File.CreateText(Path.Combine(_logPath, timestamp.ToString("yyyy-MM-dd_HH-mm-ss") + ".txt")))
                {
                    logWriter.WriteLine(exception.Message);
                }
            }

            return points;
        }
    }
}