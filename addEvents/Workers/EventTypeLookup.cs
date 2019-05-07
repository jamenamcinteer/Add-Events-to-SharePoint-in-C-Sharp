using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.SharePoint.Client;
using OfficeOpenXml;
using addEvents.Data;

namespace addEvents.Workers
{
    public class EventTypeLookup
    {
        public List<EventType> EventTypes = new List<EventType>();
        public EventTypeLookup(Web rootweb, ClientContext SAcontext, string url)
        {
            List libEventTypes = rootweb.Lists.GetByTitle("Event Types");

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><Query><Where>" +
                "</Where></Query><RowLimit>100</RowLimit></View>";
            ListItemCollection collListItem = libEventTypes.GetItems(camlQuery);

            SAcontext.Load(collListItem);

            SAcontext.ExecuteQuery();

            foreach (ListItem oListItem in collListItem)
            {
                EventType newEventType = new EventType(oListItem.Id, oListItem["Title"].ToString());
                EventTypes.Add(newEventType);
            }
        }
    }
}
