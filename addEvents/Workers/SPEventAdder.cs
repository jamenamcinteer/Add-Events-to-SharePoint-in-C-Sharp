using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using addEvents.Workers;
using addEvents.Data;

namespace addEvents.Workers
{
    class SPEventAdder
    {
        public SPEventAdder(List<Event> eventList, string url, string doclib)
        {
            foreach (Event newEvent in eventList)
            {
                AddEvent(url, doclib, newEvent);
            }
        }
        private void AddEvent(string url, string doclib, Event newEvent)
        {
            using (ClientContext SAcontext = new ClientContext(url))
            {
                SAcontext.Credentials = Helper.credential.getCredential();
                Web rootweb = SAcontext.Web;
                SAcontext.Load(rootweb, r => r.ServerRelativeUrl, r => r.AllProperties);
                List lib = rootweb.Lists.GetByTitle(doclib);
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newListItem = lib.AddItem(itemCreateInfo);
                newListItem["Title"] = newEvent.Title;
                newListItem["Date"] = newEvent.Date;
                newListItem["Start_x0020_Date_x002d_Time"] = newEvent.StartDateTime;
                newListItem["End_x0020_Date_x002d_Time"] = newEvent.EndDateTime;
                newListItem["Calendar_x0020_Ordering_x0020_Da"] = $"{newEvent.CalendarOrderingDate} {newEvent.StartTime}";
                newListItem["Time"] = newEvent.Time;
                newListItem["Region"] = newEvent.Region;
                newListItem["Short_x0020_Description"] = newEvent.ShortDescription;
                newListItem["Long_x0020_Description"] = newEvent.LongDescription;
                newListItem["Location"] = newEvent.Location;
                FieldUrlValue mapLink = new FieldUrlValue();
                mapLink.Description = newEvent.MapDescription;
                mapLink.Url = newEvent.MapUrl;
                newListItem["Map_x0020_Link"] = mapLink;
                newListItem["Cost"] = newEvent.Cost;
                newListItem["Contact_x0020_Name"] = newEvent.ContactName;
                newListItem["Contact_x0020_Phone"] = newEvent.ContactPhone;

                newListItem["Event_x0020_Type"] = newEvent.GetEventTypeLookupValues();

                newListItem["Sites"] = newEvent.GetEventSiteLookupValues();

                newListItem.Update();
                SAcontext.Load(newListItem);
                SAcontext.ExecuteQuery();

                Logger log = new Logger();
                log.Log($"Event Created: {newListItem["Title"]}", ConsoleColor.Green);
            }
        }
    }
}
