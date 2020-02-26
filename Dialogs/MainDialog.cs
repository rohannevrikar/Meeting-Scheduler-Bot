// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Azure.Cosmos.Table;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using TeamsAuth;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Graph;
using Microsoft.Bot.Builder.Dialogs.Choices;

namespace Microsoft.BotBuilderSamples
{
    public class MainDialog : LogoutDialog
    {
        protected readonly ILogger Logger;
        private readonly IConfiguration configuration;
        private string userEmail;
        private CloudTable table;
        List<TimeSlot> timeSuggestions = null;
        public MainDialog(IConfiguration configuration, ILogger<MainDialog> logger)
            : base(nameof(MainDialog), configuration["ConnectionName"])
        {
            Logger = logger;
            this.configuration = configuration;
            
            AddDialog(new OAuthPrompt(
                nameof(OAuthPrompt),
                new OAuthPromptSettings
                {
                    ConnectionName = ConnectionName,
                    Text = "Please Sign In",
                    Title = "Sign In",
                    Timeout = 300000, // User has 5 minutes to login (1000 * 60 * 5)
                }));
            
            AddDialog(new TextPrompt(nameof(TextPrompt)));
            AddDialog(new ConfirmPrompt(nameof(ConfirmPrompt)));
            AddDialog(new ChoicePrompt(nameof(ChoicePrompt)));
            AddDialog(new DateTimePrompt(nameof(DateTimePrompt)));

            AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[]
            {
                PromptStepAsync,
                AskForAttendees,
                GetTokenWithTextResultAsync,
                AskForDuration,
                GetTokenAsync,
                ShowMeetingTimeSuggestions,
                AskForTitle,
                AskForDescription,
                GetTokenWithTextResultAsync,
                SendMeetingInvite
            }));

            // The initial child Dialog to run.
            InitialDialogId = nameof(WaterfallDialog);
            
        }

        #region Bot flow methods
        private async Task<DialogTurnResult> GetTokenWithTextResultAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var result = (string)stepContext.Result;
            if (result != null)
            {
                stepContext.Context.TurnState.Add("data", result);
                return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
            }
            await stepContext.Context.SendActivityAsync("Something went wrong. Please type anything to get started again.");
            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);


        }
        private async Task<DialogTurnResult> PromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
         
            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
        }
        private async Task<DialogTurnResult> AskForAttendees(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Get the token from the previous step. Note that we could also have gotten the
            // token directly from the prompt itself. There is an example of this in the next method.
            var tokenResponse = (TokenResponse)stepContext.Result;
            if (tokenResponse?.Token != null)
            {
                // Pull in the data from the Microsoft Graph.
                var client = new SimpleGraphClient(tokenResponse.Token);
                var me = await client.GetMeAsync();
                userEmail = me.UserPrincipalName;
                table = CreateTableAsync("botdata");
                MeetingDetail meetingDetail = new MeetingDetail(me.UserPrincipalName);
                await InsertOrMergeEntityAsync(table, meetingDetail);

                return await stepContext.PromptAsync(nameof(TextPrompt), new PromptOptions { Prompt = MessageFactory.Text("With whom would you like to set up a meeting?") }, cancellationToken);
            }

            await stepContext.Context.SendActivityAsync("Something went wrong. Please type anything to get started again.");
            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }
        private async Task<DialogTurnResult> AskForDuration(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            string result = (string)stepContext.Context.TurnState["data"];

            var tokenResponse = (TokenResponse)stepContext.Result;
            if (tokenResponse?.Token != null)
            {
                var client = new SimpleGraphClient(tokenResponse.Token);
                string[] attendeeNames = string.Concat(result.Where(c => !char.IsWhiteSpace(c))).Split(",");
              
                List<string> attendeeTableStorage = new List<string>();
                foreach (string name in attendeeNames)
                {
                    List<string> attendeeEmails = await client.GetAttendeeEmailFromName(name);
                    if(attendeeEmails.Count > 1)
                    {

                        await stepContext.Context.SendActivityAsync("There are " + attendeeEmails.Count + " people whose name start with " + name + ". Please type hi to start again, and instead of first name, enter email to avoid ambiguity.");
                        return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);


                    }
                    else if (attendeeEmails.Count == 1)
                    {
                        attendeeTableStorage.Add(attendeeEmails[0]);
                    }
                    else
                    {
                        await stepContext.Context.SendActivityAsync("Attendee not found, please type anything to start again");
                        return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
                    }
                       

                }


                var sb = new System.Text.StringBuilder();
                foreach (string email in attendeeTableStorage)
                {
                    sb.Append(email + ",");
                }
                string finalString = sb.ToString().Remove(sb.Length - 1);
                if (result != null)
                {
                    MeetingDetail meetingDetail = new MeetingDetail(userEmail);
                    meetingDetail.Attendees = finalString;

                    await InsertOrMergeEntityAsync(table, meetingDetail);

                    return await stepContext.PromptAsync(nameof(TextPrompt), new PromptOptions { Prompt = MessageFactory.Text("What will be duration of the meeting? (in hours)") }, cancellationToken);
                }
            }
            await stepContext.Context.SendActivityAsync("Something went wrong. Please type anything to get started again.");

            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }
        private async Task<DialogTurnResult> GetTokenAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var result = (string)stepContext.Result;
            if (result != null)
            {
                MeetingDetail meetingDetail = new MeetingDetail(userEmail);
                meetingDetail.Duration = result;

                await InsertOrMergeEntityAsync(table, meetingDetail);
                return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
            }
            await stepContext.Context.SendActivityAsync("Something went wrong. Please type anything to get started again.");
            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);

        }

        private async Task<DialogTurnResult> ShowMeetingTimeSuggestions(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var tokenResponse = (TokenResponse)stepContext.Result;
            if (tokenResponse?.Token != null)
            {
                // Pull in the data from the Microsoft Graph.
                var client = new SimpleGraphClient(tokenResponse.Token);
                MeetingDetail meetingDetail = await RetrieveMeetingDetailsAsync(table, userEmail, userEmail);
                if (meetingDetail == null)
                {
                    await stepContext.Context.SendActivityAsync("meeting details null");
                }
                timeSuggestions = await client.GetFindMeetingTimes(meetingDetail.Attendees, Convert.ToDouble(meetingDetail.Duration));
                if (timeSuggestions.Count == 0)
                {
                    await stepContext.Context.SendActivityAsync("No appropriate meeting slot found. Please try again by typing 'hi' and change date this time");
                    return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);

                }
                var cardOptions = new List<Choice>();
                for (int i = 0; i < timeSuggestions.Count; i++)
                {
                    cardOptions.Add(new Choice() { Value = timeSuggestions[i].Start.DateTime + " - " + timeSuggestions[i].End.DateTime });
                }
                var options = new PromptOptions()
                {
                    Prompt = MessageFactory.Text("These are the time suggestions"),
                    RetryPrompt = MessageFactory.Text("Please choose an appropriate option"),
                    Choices = cardOptions,
                };

                return await stepContext.PromptAsync(nameof(ChoicePrompt), options, cancellationToken);

            }
            await stepContext.Context.SendActivityAsync("Something went wrong. Please type anything to get started again.");
            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }

        private async Task<DialogTurnResult> AskForTitle(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var result = (FoundChoice)stepContext.Result;
           
            if (result != null)
            {
                MeetingDetail meetingDetail = new MeetingDetail(userEmail);
                meetingDetail.TimeSlotChoice = result.Index.ToString();

                await InsertOrMergeEntityAsync(table, meetingDetail);

                return await stepContext.PromptAsync(nameof(TextPrompt), new PromptOptions { Prompt = MessageFactory.Text("Please enter title of the meeting")}, cancellationToken);

            }
            await stepContext.Context.SendActivityAsync("Something went wrong. Please type anything to get started again.");

            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }
        private async Task<DialogTurnResult> AskForDescription(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var result = (string)stepContext.Result;

            if (result != null)
            {
                MeetingDetail meetingDetail = new MeetingDetail(userEmail);
                meetingDetail.Title = result;

                await InsertOrMergeEntityAsync(table, meetingDetail);

                return await stepContext.PromptAsync(nameof(TextPrompt), new PromptOptions { Prompt = MessageFactory.Text("Please enter description of the meeting") }, cancellationToken);

            }
            await stepContext.Context.SendActivityAsync("Something went wrong. Please type anything to get started again.");

            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }

        private async Task<DialogTurnResult> SendMeetingInvite(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            string description = (string)stepContext.Context.TurnState["data"];
            var tokenResponse = (TokenResponse)stepContext.Result;
            if (tokenResponse?.Token != null)
            {
                var client = new SimpleGraphClient(tokenResponse.Token);
                MeetingDetail meetingDetail = await RetrieveMeetingDetailsAsync(table, userEmail, userEmail);

                await client.SendMeetingInviteAsync(timeSuggestions[Int32.Parse(meetingDetail.TimeSlotChoice)], meetingDetail.Attendees, meetingDetail.Title, description);

                await stepContext.Context.SendActivityAsync("Meeting has been scheduled. Thank you!");


            }
            else
                await stepContext.Context.SendActivityAsync("Something went wrong. Please type anything to get started again.");
            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }
        #endregion
        #region Table storage methods
        public CloudTable CreateTableAsync(string tableName)
        {
            Microsoft.Azure.Cosmos.Table.CloudStorageAccount storageAccount = Microsoft.Azure.Cosmos.Table.CloudStorageAccount.Parse(
    configuration["StorageConnectionString"]);
            // Create the queue client.

            // Create a table client for interacting with the table service
            CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());


            // Create a table client for interacting with the table service 
            CloudTable table = tableClient.GetTableReference(tableName);
            table.CreateIfNotExists();
            return table;
        }
        public static async Task<MeetingDetail> InsertOrMergeEntityAsync(CloudTable table, MeetingDetail entity)
        {
            if (entity == null)
            {
                throw new ArgumentNullException("entity");
            }
            try
            {
                // Create the InsertOrReplace table operation
                TableOperation insertOrMergeOperation = TableOperation.InsertOrMerge(entity);

                // Execute the operation.
                TableResult result = await table.ExecuteAsync(insertOrMergeOperation);
                MeetingDetail insertedCustomer = result.Result as MeetingDetail;

                // Get the request units consumed by the current operation. RequestCharge of a TableResult is only applied to Azure Cosmos DB


                return insertedCustomer;
            }
            catch (Microsoft.Azure.Cosmos.Table.StorageException e)
            {
                Console.WriteLine(e.Message);
                Console.ReadLine();
                throw;
            }
        }
        public static async Task<MeetingDetail> RetrieveMeetingDetailsAsync(CloudTable table, string partitionKey, string rowKey)
        {
            try
            {
                TableOperation retrieveOperation = TableOperation.Retrieve<MeetingDetail>(partitionKey, rowKey);
                TableResult result = await table.ExecuteAsync(retrieveOperation);
                MeetingDetail meetingDetail = result.Result as MeetingDetail;
                if (meetingDetail != null)
                {
                    return meetingDetail;
                }

                return null;

               
            }
            catch (StorageException e)
            {
                Console.WriteLine(e.Message);
                Console.ReadLine();
                throw;
            }
        }
        #endregion
    }
}
