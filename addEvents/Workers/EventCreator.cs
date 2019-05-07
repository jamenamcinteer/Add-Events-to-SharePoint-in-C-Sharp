using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using addEvents.Data;

namespace addEvents.Workers
{
    class EventCreator
    {
        private int row;
        private ExcelWorksheet worksheet;
        private StreamWriter sw;
        public List<Event> eventList = new List<Event>();
        public string logfile = $@"C:\temp\log-{DateTime.Now.ToString("yyyyMMddHHmm")}.txt";
        public int originalNumberOfEvents = 0;

        public EventCreator(FileInfo excelFile, string siteUrl)
        {
            try
            {
                sw = new StreamWriter(logfile);
                using (ExcelPackage package = new ExcelPackage(excelFile))
                {
                    worksheet = package.Workbook.Worksheets[1];
                    int colCount = worksheet.Dimension.End.Column;
                    int rowCount = worksheet.Dimension.End.Row;
                    originalNumberOfEvents = rowCount - 1;
                    EventTypeLookup eventTypes;
                    EventSiteLookup eventSites;


                    using (ClientContext SAcontext = new ClientContext(siteUrl))
                    {
                        SAcontext.Credentials = Helper.credential.getCredential();
                        Web rootweb = SAcontext.Web;
                        SAcontext.Load(rootweb, r => r.ServerRelativeUrl, r => r.AllProperties);
                        eventTypes = new EventTypeLookup(rootweb, SAcontext, siteUrl);
                        eventSites = new EventSiteLookup(rootweb, SAcontext, siteUrl);
                    }

                    for (row = 1; row <= rowCount; row++)
                    {
                        if (row > 1)
                        {
                            try
                            {
                                Event newEvent = new Event(
                                    GetCellValue(8),
                                    GetCellValue(13),
                                    GetCellValue(21),
                                    GetCellValue(20),
                                    GetCellValue(3),
                                    GetCellValue(4),
                                    GetCellValue(7),
                                    GetCellValue(5),
                                    GetCellValue(6),
                                    GetCellValue(19),
                                    GetCellValue(9),
                                    GetCellValue(15),
                                    GetCellValue(18),
                                    GetCellValue(11),
                                    eventTypes,
                                    GetCellValue(12),
                                    eventSites,
                                    GetCellValue(2)
                                    );

                                Logger log = new Logger();
                                log.LogConsoleAndFile("Row: " + row, sw);
                                log.LogConsoleAndFile("Title: " + newEvent.Title, sw);
                                log.LogConsoleAndFile("Date(s): " + newEvent.Date, sw);
                                log.LogConsoleAndFile($"Calendar Ordering Date: {newEvent.CalendarOrderingDate} {newEvent.StartTime}", sw);
                                log.LogConsoleAndFile("Start Date: " + newEvent.StartDateTime, sw);
                                log.LogConsoleAndFile("End Date: " + newEvent.EndDateTime, sw);
                                log.LogConsoleAndFile("Time: " + newEvent.Time, sw);
                                log.LogConsoleAndFile("Location: " + newEvent.Location, sw);
                                log.LogConsoleAndFile("Region: " + newEvent.Region, sw);
                                log.LogConsoleAndFile("Map Url: " + newEvent.MapUrl, sw);
                                log.LogConsoleAndFile("Map Description: " + newEvent.MapDescription, sw);
                                log.LogConsoleAndFile("Short Description: " + newEvent.ShortDescription, sw);
                                log.LogConsoleAndFile("Long Description: " + newEvent.LongDescription, sw);
                                log.LogConsoleAndFile("Cost: " + newEvent.Cost, sw);
                                log.LogConsoleAndFile("Contact Name: " + newEvent.ContactName, sw);
                                log.LogConsoleAndFile("Contact Phone: " + newEvent.ContactPhone, sw);
                                log.LogConsoleAndFile("", sw);

                                eventList.Add(newEvent);
                            }
                            catch (Exception ex)
                            {
                                sw.WriteLine($"Error in Row {row}: {ex.Message}");
                                Logger log = new Logger();
                                log.Log($"Error in Row {row}: {ex.Message}", ConsoleColor.Red);
                                log.LogConsoleAndFile("", sw);
                            }
                        }
                    }
                }
                sw.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.Message);
                Console.WriteLine("");
            }
        }

        private object GetCellValue(int col)
        {
            return worksheet.Cells[row, col].Value;
        }
    }
}
