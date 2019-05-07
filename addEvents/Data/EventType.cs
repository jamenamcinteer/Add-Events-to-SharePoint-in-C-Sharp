using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace addEvents.Data
{
    public class EventType
    {
        public int ID { get; set; }
        public string Title { get; set; }

        public EventType(int id, string title)
        {
            ID = id;
            Title = title;
        }
    }
}
