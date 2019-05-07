using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using Microsoft.SharePoint.Client;
using System.Net;
using OfficeOpenXml;
using addEvents.Workers;
using addEvents.Data;

namespace addEvents
{
    class site
    {
        public string id { get; set; }
        public string url { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string doclib = "Events";
                List<EventType> eventTypes = new List<EventType>();
                List<EventSite> eventSites = new List<EventSite>();

                Console.WriteLine("Enter the full path and file name of the Excel document containing events to be added to SharePoint:");
                string excelFileName = Console.ReadLine();
                FileInfo excelFile = new FileInfo(excelFileName);

                Console.WriteLine("Enter the SharePoint site url:");
                string siteUrl = Console.ReadLine() ?? string.Empty;

                EventCreator ec = new EventCreator(excelFile, siteUrl);
                Console.WriteLine($"Number of events: {ec.eventList.Count} (Errors in {ec.originalNumberOfEvents - ec.eventList.Count})");
                Console.WriteLine("Log file created: " + ec.logfile);

                Console.Write("Would you like to continue to add events to the site (Y/N)? : ");
                bool continueAddEvents = Console.ReadLine().ToUpper() == "Y" ? true : false;
                if (continueAddEvents)
                {
                    if (!siteUrl.Equals(string.Empty))
                    {
                        SPEventAdder eventAdder = new SPEventAdder(ec.eventList, siteUrl, doclib);
                    }
                }
                Console.WriteLine("Complete. Enter any key to exit...");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.Message);
                Console.WriteLine();
            }
        }
    }
}
