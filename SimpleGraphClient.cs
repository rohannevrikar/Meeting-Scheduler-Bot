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

        // Sends an email on the users behalf using the Microsoft Graph API
        public async Task SendMailAsync(string toAddress, string subject, string content)
        {
            if (string.IsNullOrWhiteSpace(toAddress))
            {
                throw new ArgumentNullException(nameof(toAddress));
            }

            if (string.IsNullOrWhiteSpace(subject))
            {
                throw new ArgumentNullException(nameof(subject));
            }

            if (string.IsNullOrWhiteSpace(content))
            {
                throw new ArgumentNullException(nameof(content));
            }

            var graphClient = GetAuthenticatedClient();
            var recipients = new List<Recipient>
            {
                new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = toAddress,
                    },
                },
            };

            // Create the message.
            var email = new Message
            {
                Body = new ItemBody
                {
                    Content = content,
                    ContentType = BodyType.Text,
                },
                Subject = subject,
                ToRecipients = recipients,
            };

            // Send the message.
            await graphClient.Me.SendMail(email, true).Request().PostAsync();
        }
        public async Task SendMeetingInviteAsync(TimeSlot timeSlot, string attendees)
        {
            var graphClient = GetAuthenticatedClient();
            List<string> attendeeEmails = await GetAttendeesEmails(attendees);
            var calendar = await graphClient.Me.Calendar
    .Request()
    .GetAsync();
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
                Subject = "Meeting invite",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = "Does mid month work for you?"
                },
                Start = new DateTimeTimeZone
                {
                    DateTime = timeSlot.Start.DateTime,
                    TimeZone = "Pacific Standard Time"
                },
                End = new DateTimeTimeZone
                {
                    DateTime = timeSlot.End.DateTime,
                    TimeZone = "Pacific Standard Time"
                },
             
                Attendees = attendeeList
            };

            await graphClient.Me.Calendars[calendar.Id].Events
                .Request()
                .AddAsync(@event);

        }
        // Gets mail for the user using the Microsoft Graph API
        public async Task<Message[]> GetRecentMailAsync()
        {

            var graphClient = GetAuthenticatedClient();
            var messages = await graphClient.Me.MailFolders.Inbox.Messages.Request().GetAsync();
            return messages.Take(5).ToArray();
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

                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);
                var body = new 
                {
                    attendeeList,
                    meetingDuration
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
       

        // Get information about the user.
        public async Task<User> GetMeAsync()
        {
            var graphClient = GetAuthenticatedClient();
            var me = await graphClient.Me.Request().GetAsync();
            return me;
        }

        // gets information about the user's manager.
        public async Task<User> GetManagerAsync()
        {
            var graphClient = GetAuthenticatedClient();
            var manager = await graphClient.Me.Manager.Request().GetAsync() as User;
            return manager;
        }

        // // Gets the user's photo
        // public async Task<PhotoResponse> GetPhotoAsync()
        // {
        //     HttpClient client = new HttpClient();
        //     client.DefaultRequestHeaders.Add("Authorization", "Bearer " + _token);
        //     client.DefaultRequestHeaders.Add("Accept", "application/json");

        //     using (var response = await client.GetAsync("https://graph.microsoft.com/v1.0/me/photo/$value"))
        //     {
        //         if (!response.IsSuccessStatusCode)
        //         {
        //             throw new HttpRequestException($"Graph returned an invalid success code: {response.StatusCode}");
        //         }

        //         var stream = await response.Content.ReadAsStreamAsync();
        //         var bytes = new byte[stream.Length];
        //         stream.Read(bytes, 0, (int)stream.Length);

        //         var photoResponse = new PhotoResponse
        //         {
        //             Bytes = bytes,
        //             ContentType = response.Content.Headers.ContentType?.ToString(),
        //         };

        //         if (photoResponse != null)
        //         {
        //             photoResponse.Base64String = $"data:{photoResponse.ContentType};base64," +
        //                                          Convert.ToBase64String(photoResponse.Bytes);
        //         }

        //         return photoResponse;
        //     }
        // }

        // Get an Authenticated Microsoft Graph client using the token issued to the user.
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
