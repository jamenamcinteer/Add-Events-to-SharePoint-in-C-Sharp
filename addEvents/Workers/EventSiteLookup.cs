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
    public class EventSiteLookup
    {
        public List<EventSite> EventSites = new List<EventSite>();
        public EventSiteLookup(Web rootweb, ClientContext SAcontext, string url)
        {
            List libEventTypes = rootweb.Lists.GetByTitle("Sites");

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><Query><Where>" +
                "</Where></Query><RowLimit>100</RowLimit></View>";
            ListItemCollection collListItem = libEventTypes.GetItems(camlQuery);

            SAcontext.Load(collListItem);

            SAcontext.ExecuteQuery();

            foreach (ListItem oListItem in collListItem)
            {
                EventSite newEventSite = new EventSite(oListItem.Id, oListItem["Title"].ToString());
                EventSites.Add(newEventSite);
            }
        }
    }
}
