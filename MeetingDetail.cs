using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Azure.Cosmos.Table;

namespace TeamsAuth
{
    public class MeetingDetail : TableEntity
    {
        public MeetingDetail()
        {
        }

        public MeetingDetail(string userEmail)
        {
            PartitionKey = userEmail;
            RowKey = userEmail;
        }
        public string Title { get; set; }
        public string TimeSlotChoice { get; set; }

        public string Attendees { get; set; }
        public string StartDateTime { get; set; }
        public string Duration { get; set; }

    }
}
