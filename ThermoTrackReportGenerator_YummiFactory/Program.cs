using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace ThermoTrackReportGenerator_YummiFactory
{
    class Program
    {
        static void Main(string[] args)
        {
            string documentName = "YummiFactory_Temperatures_Week_16";
            string documentPath = @"D:\LUDIG\DOCUMENTS\THERMOTRACK\CLIENTS\YummiFactory\Reports\Week_16\";

            Document document = new Document();

            {
                int indexSection = document.AddSection();

                {
                    DateTime startTime = new DateTime(2019, 4, 12, 0, 0, 0);
                    DateTime endTime = new DateTime(2019, 4, 19, 0, 0, 0);

                    {
                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Centre);
                        document.AddRun(indexSection, indexElement, "Yummi Factory - Paarl", true, true, 70);
                        document.AddRun(indexSection, indexElement, true);
                        document.AddRun(indexSection, indexElement, "Sensor Readings - Temperatures", true, true, 50);
                        document.AddRun(indexSection, indexElement, true);
                        document.AddRun(indexSection, indexElement, startTime.ToLongDateString() + " - " + endTime.ToLongDateString() + " (Not including)", true, true, 40);
                    }

                    {
                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Left);
                        document.AddRun(indexSection, indexElement, true);
                        document.AddRun(indexSection, indexElement, true);
                        document.AddRun(indexSection, indexElement, "The following charts represent the sensor readings for the ");
                        document.AddRun(indexSection, indexElement, "Yummi Factory", false, true);
                        document.AddRun(indexSection, indexElement, " in ");
                        document.AddRun(indexSection, indexElement, "Paarl", false, true);
                        document.AddRun(indexSection, indexElement, ", concerning ");
                        document.AddRun(indexSection, indexElement, "temperature", false, true);
                        document.AddRun(indexSection, indexElement, " as recorded over the period of ");
                        document.AddRun(indexSection, indexElement, startTime.ToLongDateString(), false, true);
                        document.AddRun(indexSection, indexElement, " up to (and ");
                        document.AddRun(indexSection, indexElement, "NOT", false, true);
                        document.AddRun(indexSection, indexElement, " including) ");
                        document.AddRun(indexSection, indexElement, endTime.ToLongDateString(), false, true);
                        document.AddRun(indexSection, indexElement, ".");
                    }

                    {
                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Left);
                        document.AddRun(indexSection, indexElement, true);
                        document.AddRun(indexSection, indexElement, "Sensor-Zone Mappings", true, true);
                        document.AddRun(indexSection, indexElement, true);
                        document.AddRun(indexSection, indexElement, true);
                        document.AddRun(indexSection, indexElement, "The sensors-to-zone mappings are as follows:");
                    }

                    {
                        {
                            int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Left);
                            document.AddRun(indexSection, indexElement, true);
                            document.AddRun(indexSection, indexElement, "Above-Zero Cold Room:", false, true);
                            document.AddRun(indexSection, indexElement, true);
                        }

                        {
                            int idListFormat = document.AddListFormat();

                            {
                                int indexElement = document.AddListItem(indexSection, idListFormat);
                                document.AddRun(indexSection, indexElement, "0000001E0034");
                            }

                            {
                                int indexElement = document.AddListItem(indexSection, idListFormat);
                                document.AddRun(indexSection, indexElement, "00000021002D");
                            }

                            {
                                int indexElement = document.AddListItem(indexSection, idListFormat);
                                document.AddRun(indexSection, indexElement, "00000022002D");
                            }
                        }
                    }

                    {
                        {
                            int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Left);
                            document.AddRun(indexSection, indexElement, true);
                            document.AddRun(indexSection, indexElement, "Sub-Zero Cold Room:", false, true);
                            document.AddRun(indexSection, indexElement, true);
                        }

                        {
                            int idListFormat = document.AddListFormat();

                            {
                                int indexElement = document.AddListItem(indexSection, idListFormat);
                                document.AddRun(indexSection, indexElement, "0000001E0038");
                            }

                            {
                                int indexElement = document.AddListItem(indexSection, idListFormat);
                                document.AddRun(indexSection, indexElement, "0000001E0039");
                            }

                            {
                                int indexElement = document.AddListItem(indexSection, idListFormat);
                                document.AddRun(indexSection, indexElement, "0000001F004C");
                            }
                        }
                    }

                    {
                        {
                            int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Left);
                            document.AddRun(indexSection, indexElement, true);
                            document.AddRun(indexSection, indexElement, "Kitchen Area:", false, true);
                            document.AddRun(indexSection, indexElement, true);
                        }

                        {
                            int idListFormat = document.AddListFormat();

                            {
                                int indexElement = document.AddListItem(indexSection, idListFormat);
                                document.AddRun(indexSection, indexElement, "000000220043");
                            }
                        }
                    }
                }
            }

            {
                int indexSection = document.AddSection();

                {
                    {
                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Centre);
                        document.AddRun(indexSection, indexElement, "Above-Zero Cold Room:", true, true, 60);
                    }

                    {
                        string beaconID = "0000001E0034";

                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Centre);
                        document.AddRun(indexSection, indexElement, beaconID, true, true, 50);

                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 13, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 13, 0, 0, 0), new DateTime(2019, 4, 14, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 14, 0, 0, 0), new DateTime(2019, 4, 15, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 15, 0, 0, 0), new DateTime(2019, 4, 16, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 16, 0, 0, 0), new DateTime(2019, 4, 17, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 17, 0, 0, 0), new DateTime(2019, 4, 18, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 18, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                    }
                }
            }

            {
                int indexSection = document.AddSection();

                {
                    {
                        string beaconID = "00000021002D";

                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Centre);
                        document.AddRun(indexSection, indexElement, beaconID, true, true, 50);

                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 13, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 13, 0, 0, 0), new DateTime(2019, 4, 14, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 14, 0, 0, 0), new DateTime(2019, 4, 15, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 15, 0, 0, 0), new DateTime(2019, 4, 16, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 16, 0, 0, 0), new DateTime(2019, 4, 17, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 17, 0, 0, 0), new DateTime(2019, 4, 18, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 18, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                    }
                }
            }

            {
                int indexSection = document.AddSection();

                {
                    {
                        string beaconID = "00000022002D";

                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Centre);
                        document.AddRun(indexSection, indexElement, beaconID, true, true, 50);

                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 13, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 13, 0, 0, 0), new DateTime(2019, 4, 14, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 14, 0, 0, 0), new DateTime(2019, 4, 15, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 15, 0, 0, 0), new DateTime(2019, 4, 16, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 16, 0, 0, 0), new DateTime(2019, 4, 17, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 17, 0, 0, 0), new DateTime(2019, 4, 18, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 18, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                    }
                }
            }

            {
                int indexSection = document.AddSection();

                {
                    {
                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Centre);
                        document.AddRun(indexSection, indexElement, "Sub-Zero Cold Room:", true, true, 60);
                    }

                    {
                        string beaconID = "0000001E0038";

                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Centre);
                        document.AddRun(indexSection, indexElement, beaconID, true, true, 50);

                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 13, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 13, 0, 0, 0), new DateTime(2019, 4, 14, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 14, 0, 0, 0), new DateTime(2019, 4, 15, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 15, 0, 0, 0), new DateTime(2019, 4, 16, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 16, 0, 0, 0), new DateTime(2019, 4, 17, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 17, 0, 0, 0), new DateTime(2019, 4, 18, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 18, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                    }
                }
            }

            {
                int indexSection = document.AddSection();

                {
                    {
                        string beaconID = "0000001E0039";

                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Centre);
                        document.AddRun(indexSection, indexElement, beaconID, true, true, 50);

                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 13, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 13, 0, 0, 0), new DateTime(2019, 4, 14, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 14, 0, 0, 0), new DateTime(2019, 4, 15, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 15, 0, 0, 0), new DateTime(2019, 4, 16, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 16, 0, 0, 0), new DateTime(2019, 4, 17, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 17, 0, 0, 0), new DateTime(2019, 4, 18, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 18, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                    }
                }
            }

            {
                int indexSection = document.AddSection();

                {
                    {
                        string beaconID = "0000001F004C";

                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Centre);
                        document.AddRun(indexSection, indexElement, beaconID, true, true, 50);

                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 13, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 13, 0, 0, 0), new DateTime(2019, 4, 14, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 14, 0, 0, 0), new DateTime(2019, 4, 15, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 15, 0, 0, 0), new DateTime(2019, 4, 16, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 16, 0, 0, 0), new DateTime(2019, 4, 17, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 17, 0, 0, 0), new DateTime(2019, 4, 18, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 18, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                    }
                }
            }

            {
                int indexSection = document.AddSection();

                {
                    {
                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Centre);
                        document.AddRun(indexSection, indexElement, "Kitchen Area:", true, true, 60);
                    }

                    {
                        string beaconID = "000000220043";

                        int indexElement = document.AddParagraph(indexSection, Document.AlignmentHorizontal.Centre);
                        document.AddRun(indexSection, indexElement, beaconID, true, true, 50);

                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 12, 0, 0, 0), new DateTime(2019, 4, 13, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 13, 0, 0, 0), new DateTime(2019, 4, 14, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 14, 0, 0, 0), new DateTime(2019, 4, 15, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 15, 0, 0, 0), new DateTime(2019, 4, 16, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 16, 0, 0, 0), new DateTime(2019, 4, 17, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 17, 0, 0, 0), new DateTime(2019, 4, 18, 0, 0, 0));
                        AddTemperaturesByTimeGraph(ref document, indexSection, beaconID, new DateTime(2019, 4, 18, 0, 0, 0), new DateTime(2019, 4, 19, 0, 0, 0));
                    }
                }
            }

            if (!document.CloseAndWriteToFile(documentPath + documentName + ".docx"))
                Debug.WriteLine("Failed to write the document file to disk");
        }

        private static void AddTemperaturesByTimeGraph(ref Document document, int indexSection, string beaconID, DateTime startTime, DateTime endTime)
        {
            MySQLController mySQLController = new MySQLController();

            DateTime timestamp = DateTime.Now;

            List<KeyValuePair<DateTime, float>> temperaturesByTime = mySQLController.GetTemperaturesByTime(beaconID, startTime, endTime, timestamp);

            List<KeyValuePair<double, double>> points = new List<KeyValuePair<double, double>>();

            foreach (KeyValuePair<DateTime, float> temperatureByTime in temperaturesByTime)
                points.Add(new KeyValuePair<double, double>(temperatureByTime.Key.ToOADate(), Math.Round(temperatureByTime.Value, 1)));

            string chartName = beaconID + ": Temperatures from " + startTime.ToString("yyyy-MM-dd") + " to " + endTime.ToString("yyyy-MM-dd");

            List<Series_Point> series = new List<Series_Point>();
            series.Add(new Series_Point(beaconID, points, 0));

            document.AddScatterChart(indexSection, chartName, series, true, false, false, startTime.ToOADate(), endTime.ToOADate(), Document.UnitType.None, Document.UnitType.Temperature_Celsius, "Time", "Temperature");
        }

        private static DataSet ReadExcelFile(string filePath)
        {
            DataSet set = new DataSet();

            string connectionString = GetConnectionString(filePath);

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;

                // Get all Sheets in Excel File.

                DataTable tableSheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data.

                foreach (DataRow row in tableSheets.Rows)
                {
                    string sheetName = (row["TABLE_NAME"].ToString()).Trim('\'');

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet

                    command.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable table = new DataTable();
                    table.TableName = sheetName.TrimEnd('$');

                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                    adapter.Fill(table);

                    set.Tables.Add(table);
                }

                command = null;
                connection.Close();
            }

            return set;
        }

        private static string GetConnectionString(string filePath)
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013.

            properties["Provider"] = "Microsoft.ACE.OLEDB.12.0";
            properties["Data Source"] = filePath;
            properties["Extended Properties"] = "\"Excel 12.0 XML";

            // XLS - Excel 2003 and Older.

            //properties["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            //properties["Data Source"] = "C:\\MyExcel.xls";
            //properties["Extended Properties"] = "\"Excel 8.0";

            StringBuilder connectionStringBuilder = new StringBuilder();

            foreach (KeyValuePair<string, string> property in properties)
            {
                connectionStringBuilder.Append(property.Key);
                connectionStringBuilder.Append('=');
                connectionStringBuilder.Append(property.Value);
                connectionStringBuilder.Append(';');
            }

            connectionStringBuilder.Append("IMEX=1\";");

            return connectionStringBuilder.ToString();
        }
    }
}