// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace Microsoft.BotBuilderSamples
{
    // This class is a wrapper for the Microsoft Graph API
    // See: https://developer.microsoft.com/en-us/graph
    public class SimpleGraphClient
    {
        private readonly string _token;

        public SimpleGraphClient(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new ArgumentNullException(nameof(token));
            }

            _token = token;
        }

        public async Task SendMeetingInviteAsync(TimeSlot timeSlot, string attendees, string title, string description)
        {
            var graphClient = GetAuthenticatedClient();
            List<string> attendeeEmails = await GetAttendeesEmails(attendees);
            var calendar = await graphClient.Me.Calendar.Request().GetAsync();
            var attendeeList = new List<Attendee>();

         
            foreach (string email in attendeeEmails)
            {
                attendeeList.Add(
                new Attendee
                {
                    Type = AttendeeType.Required,
                    EmailAddress = new EmailAddress
                    {
                        Address = email
                    }
                });
            }

            var @event = new Event
            {
                Subject = title,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = description
                },
                Start = new DateTimeTimeZone
                {
                    DateTime = timeSlot.Start.DateTime,
                    TimeZone = TimeZoneInfo.Local.Id
                },
                End = new DateTimeTimeZone
                {
                    DateTime = timeSlot.End.DateTime,
                    TimeZone = TimeZoneInfo.Local.Id
                },
             
                Attendees = attendeeList
            };

            await graphClient.Me.Calendars[calendar.Id].Events
                .Request()
                .AddAsync(@event);

        }
        public async Task<List<string>> GetAttendeesEmails(string attendees)
        {
            var graphClient = GetAuthenticatedClient();
            var users = await graphClient.Users.Request().GetAsync();
            string[] attendeeNames = string.Concat(attendees.Where(c => !char.IsWhiteSpace(c))).Split(",");
            List<string> attendeeEmails = new List<string>();
            foreach(string name in attendeeNames)
            {
                try
                {
                    attendeeEmails.Add(users.Where(a => a.DisplayName.Contains(name, StringComparison.OrdinalIgnoreCase) || a.UserPrincipalName.Contains(name, StringComparison.OrdinalIgnoreCase)).Select(a => a.UserPrincipalName).First());

                }
                catch (InvalidOperationException e)
                {
                    return null;
                }
            }
            return attendeeEmails;
        }
        public async Task<List<TimeSlot>> GetFindMeetingTimes(string attendees, double duration)
        {
            try
            {
                var graphClient = GetAuthenticatedClient();
                List<string> attendeeEmails = await GetAttendeesEmails(attendees);
                var attendeeList = new List<AttendeeBase>();
                foreach (string email in attendeeEmails)
                {
                    attendeeList.Add(
                    new AttendeeBase
                    {
                        Type = AttendeeType.Required,
                        EmailAddress = new EmailAddress
                        {

                            Address = email
                        }
                    });
                }

                var meetingDuration = new Duration(System.Xml.XmlConvert.ToString(TimeSpan.FromHours(duration)));
                var minimumAttendeePercentage = 100;

                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);
                var body = new 
                {
                    attendeeList,
                    meetingDuration,
                    minimumAttendeePercentage
                };
                HttpResponseMessage response = await client.PostAsJsonAsync(
                "https://graph.microsoft.com/v1.0/me/findMeetingTimes", body);
                response.EnsureSuccessStatusCode();

                // return URI of the created resource.
                
                var meetingTimeSuggestionsResult = await response.Content.ReadAsAsync<MeetingTimeSuggestionsResult>();
                //var meetingTimeSuggestionsResult = await graphClient.Me
                //    .FindMeetingTimes(attendeeList, null, null, null)
                //    .Request()
                //    .Header("Prefer", "outlook.timezone=\"Pacific Standard Time\"")
                //    .PostAsync();


                var timeSuggestions = new List<TimeSlot>();
                foreach (MeetingTimeSuggestion meetingTimeSuggestion in meetingTimeSuggestionsResult.MeetingTimeSuggestions)
                {
                    timeSuggestions.Add(meetingTimeSuggestion.MeetingTimeSlot);
                }
                return timeSuggestions;
            }
            catch(Exception e)
            {
                throw e;
            }

           
        }      
        public async Task<User> GetMeAsync()
        {
            var graphClient = GetAuthenticatedClient();
            var me = await graphClient.Me.Request().GetAsync();
            return me;
        }       
        private GraphServiceClient GetAuthenticatedClient()
        {
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    requestMessage =>
                    {
                        // Append the access token to the request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", _token);

                        // Get event times in the current time zone.
                        requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

                        return Task.CompletedTask;
                    }));
            return graphClient;
        }
    }
}
