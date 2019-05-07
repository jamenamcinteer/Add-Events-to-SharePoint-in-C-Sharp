using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using addEvents.Workers;
using Microsoft.SharePoint.Client;

namespace addEvents.Data
{
    public class Event
    {
        private List<string> addressLinesList = new List<string>();
        private string rawLocation = string.Empty;
        private string city = string.Empty;

        public string Title { get; set; }
        public string Date { get; set; }
        public string StartDateTime { get; set; }
        public string EndDateTime { get; set; }
        public string CalendarOrderingDate { get; set; }
        public string Time { get; set; }
        public string StartTime { get; set; }
        public string Region { get; set; }
        public string ShortDescription { get; set; }
        public string LongDescription { get; set; }
        public string Location { get; set; }
        public string MapUrl { get; set; }
        public string MapDescription { get; set; }
        public string Cost { get; set; }
        public string ContactName { get; set; }
        public string ContactPhone { get; set; }
        public string EventTypeValues { get; set; }
        public EventTypeLookup EventTypes { get; set; }
        public string EventSiteValues { get; set; }
        public EventSiteLookup EventSites { get; set; }

        public Event(
            object locationCell,
            object reoccurringStartDateCell,
            object timeCell,
            object dateOfEventCell,
            object displayStartDateCell,
            object displayEndDateCell,
            object regionCell,
            object shortDescriptionCell,
            object longDescriptionCell,
            object mapLinkCell,
            object costCell,
            object contactNameCell,
            object contactPhoneCell,
            object eventTypeCell,
            EventTypeLookup eventTypes,
            object eventSiteCell,
            EventSiteLookup eventSites,
            object titleCell
        )
        {
            SetLocationValues(locationCell);
            SetTitle(titleCell);
            Date = dateOfEventCell.ToString().Trim();
            StartDateTime = displayStartDateCell.ToString().Trim();
            if (displayEndDateCell != null)
            {
                EndDateTime = displayEndDateCell.ToString().Trim();
            }
            else
            {
                throw new Exception("Display End Date is null.");
            }
            SetCalendarOrderingDate(reoccurringStartDateCell);
            SetTime(timeCell);
            Region = regionCell.ToString().Trim();
            ShortDescription = shortDescriptionCell.ToString().Trim();
            LongDescription = longDescriptionCell.ToString().Trim();
            SetLocation();
            MapUrl = mapLinkCell.ToString().Trim();
            SetMapDescription();
            Cost = costCell.ToString().Trim();
            ContactName = contactNameCell.ToString().Trim();
            ContactPhone = contactPhoneCell.ToString().Trim();
            EventTypeValues = eventTypeCell.ToString().Trim();
            EventTypes = eventTypes;
            EventSiteValues = eventSiteCell.ToString().Trim();
            EventSites = eventSites;
        }

        public FieldLookupValue[] GetEventTypeLookupValues()
        {
            string[] eventTypeCellValues = EventTypeValues.Split('/');

            FieldLookupValue[] eventTypeList = new FieldLookupValue[eventTypeCellValues.Length];

            for (int i = 0; i < eventTypeCellValues.Length; i++)
            {
                foreach (EventType eventType in EventTypes.EventTypes)
                {
                    if (eventType.Title == eventTypeCellValues[i] || eventType.Title == "General Health" && eventTypeCellValues[i] == "Preventative Treatment")
                    {
                        FieldLookupValue fieldLookupValue = new FieldLookupValue();
                        fieldLookupValue.LookupId = eventType.ID;
                        eventTypeList[i] = fieldLookupValue;
                    }
                }
            }

            return eventTypeList;
        }
        public FieldLookupValue[] GetEventSiteLookupValues()
        {
            string[] eventSiteCellValues = EventSiteValues.Split('/');

            FieldLookupValue[] eventSiteList = new FieldLookupValue[eventSiteCellValues.Length];

            for (int i = 0; i < eventSiteCellValues.Length; i++)
            {
                foreach (EventSite eventSite in EventSites.EventSites)
                {
                    if (eventSite.Title == eventSiteCellValues[i])
                    {
                        FieldLookupValue fieldLookupValue = new FieldLookupValue();
                        fieldLookupValue.LookupId = eventSite.ID;
                        eventSiteList[i] = fieldLookupValue;
                    }
                }
            }

            return eventSiteList;
        }
        private void SetLocationValues(object locationCell)
        {
            rawLocation = locationCell.ToString().Trim();
            string[] addressLines = rawLocation.IndexOf("\n") > -1 ? rawLocation.Split('\n') : rawLocation.Split(new string[] { "   " }, StringSplitOptions.None);
            foreach (string addressLine in addressLines)
            {
                if (addressLine.ToString().Trim().Length > 0)
                {
                    addressLinesList.Add(addressLine.ToString().Trim());
                }
            }
            int firstStringPosition = rawLocation.IndexOf(addressLines[addressLines.Length - 1]);
            int secondStringPosition = rawLocation.IndexOf(", NM") > -1 ? rawLocation.IndexOf(", NM") : rawLocation.IndexOf(" NM");
            city = rawLocation.Substring(firstStringPosition, secondStringPosition - firstStringPosition);
        }
        private void SetTitle(object titleCell)
        {
            Title = titleCell != null ? titleCell.ToString().Trim() : $"Event - {city}";
        }
        private void SetCalendarOrderingDate(object reoccurringStartDateCell)
        {
            string[] dates = Date.Split(',');
            CalendarOrderingDate = Date == "Monday-Friday"
                ? reoccurringStartDateCell != null ? DateTime.Parse(reoccurringStartDateCell.ToString().Trim()).ToString("MM/dd/yyyy") : DateTime.Parse(StartDateTime).ToString("MM/dd/yyyy")
                : DateTime.Parse(dates[0].ToString().Trim()).ToString("MM/dd/yyyy");
        }
        private void SetTime(object timeCell)
        {
            string unformattedTime = timeCell.ToString().Trim();
            string[] times = unformattedTime.Split('-');
            string timeStart = DateTime.Parse($"{CalendarOrderingDate} {times[0]}").ToString("h:mm tt", new DateTimeFormatInfo { AMDesignator = "a.m.", PMDesignator = "p.m." });
            string timeEnd = DateTime.Parse($"{CalendarOrderingDate} {times[1]}").ToString("h:mm tt", new DateTimeFormatInfo { AMDesignator = "a.m.", PMDesignator = "p.m." });
            StartTime = DateTime.Parse($"{CalendarOrderingDate} {times[0]}").ToString("hh:mm:ss tt", new DateTimeFormatInfo { AMDesignator = "AM", PMDesignator = "PM" });
            Time = $"{timeStart} - {timeEnd}";
        }
        private void SetLocation()
        {
            Location = string.Join("<br>", addressLinesList);
        }

        private void SetMapDescription()
        {
            MapDescription = rawLocation.IndexOf("\n") > -1 ? rawLocation.Substring(0, rawLocation.IndexOf("\n")) : rawLocation.Substring(0, rawLocation.IndexOf("   "));
        }

    }
}
