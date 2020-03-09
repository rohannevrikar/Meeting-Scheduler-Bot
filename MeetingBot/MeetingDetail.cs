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

        public MeetingDetail(string conversationId, string userEmail)
        {
            PartitionKey = userEmail;
            RowKey = conversationId;
        }
        public string Title { get; set; }
        public string Description { get; set; }

        public string TimeSlotChoice { get; set; }

        public string Attendees { get; set; }
        public string StartDateTime { get; set; }
        public double Duration { get; set; }

    }
}
