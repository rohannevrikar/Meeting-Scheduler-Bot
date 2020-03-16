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
            AddDialog(new ChoicePrompt(nameof(ChoicePrompt)) { Style = ListStyle.HeroCard });
            AddDialog(new DateTimePrompt(nameof(DateTimePrompt)));
            AddDialog(new NumberPrompt<double>(nameof(NumberPrompt<double>), DurationPromptValidatorAsync));

            AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[] //As this is a waterfall dialog, bot will interact with user in this order
            {
                GetTokenAsync, //Prompt user to sign in
                AskForAttendees, //Ask user for attendees
                GetTokenAsync, //Get token which will be used in next step to call Graph to get email of attendees and save attendees in turnState so that it can be used in next step
                AskForDuration, //Ask user for duration of the meeting
                GetTokenAsync, //Get token which will be used in next step to call Graph to get meeting times
                ShowMeetingTimeSuggestions, //Show meeting times
                AskForTitle, //Ask for title of the meeting
                AskForDescription, //Ask for description of the meeting
                GetTokenAsync, //Get token which will be used in next step to call Graph to send meeting invite and save description in turnState so that it can be used in next step
                SendMeetingInvite //Create event 
            }));

            // The initial child Dialog to run.
            InitialDialogId = nameof(WaterfallDialog);

        }

     

        #region Bot flow methods
        private async Task<DialogTurnResult> GetTokenAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if(stepContext.Result != null)
            {
               
                if (stepContext.Result.GetType().Equals(typeof(System.String)))
                {
                    stepContext.Context.TurnState.Add("data", (string)stepContext.Result); //acts as temporary intermediate storage for previous step's string input (attendees and description)
                }
                else if (stepContext.Result.GetType().Equals(typeof(System.Double)))
                {
                    stepContext.Context.TurnState.Add("data", ((double)stepContext.Result).ToString()); //acts as temporary intermediate storage for previous step's double input (double)
                }               
            }
            
            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken); //OAuthPrompt prompts user to sign in if haven't done so already and retrieves token. If the user is already signed in, then it just retrieves the token
        }
        private async Task<DialogTurnResult> AskForAttendees(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {

            var tokenResponse = (TokenResponse)stepContext.Result; // Get the token from the previous step.
            if (tokenResponse?.Token != null)
            {
                // Pull in the data from the Microsoft Graph.
                var client = new SimpleGraphClient(tokenResponse.Token);
                var me = await client.GetMeAsync();
                userEmail = me.UserPrincipalName;
                table = CreateTableAsync("botdata"); //creates table if does not exist already
                MeetingDetail meetingDetail = new MeetingDetail(stepContext.Context.Activity.Conversation.Id, userEmail);
                await InsertOrMergeEntityAsync(table, meetingDetail); //inserts user's email in table
                return await stepContext.PromptAsync(nameof(TextPrompt), new PromptOptions { Prompt = MessageFactory.Text("With whom would you like to set up a meeting?") }, cancellationToken);
            }

            await stepContext.Context.SendActivityAsync("Something went wrong. Please type anything to get started again.");
            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }
        private async Task<DialogTurnResult> AskForDuration(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            string result = (string)stepContext.Context.TurnState["data"]; //gets attendees' names

            var tokenResponse = (TokenResponse)stepContext.Result; //gets token
            if (tokenResponse?.Token != null)
            {
                var client = new SimpleGraphClient(tokenResponse.Token);
                string[] attendeeNames = string.Concat(result.Where(c => !char.IsWhiteSpace(c))).Split(","); //splits comma separated names of attendees

                List<string> attendeeTableStorage = new List<string>();
                foreach (string name in attendeeNames)
                {

                    List<string> attendeeEmails = await client.GetAttendeeEmailFromName(name); //gets email from attendee's name
                    if (attendeeEmails.Count > 1) //there can be multiple people having same first name, ask user to start again and enter email instead to be more specific
                    {

                        await stepContext.Context.SendActivityAsync("There are " + attendeeEmails.Count + " people whose name start with " + name + ". Please type hi to start again, and instead of first name, enter email to avoid ambiguity.");
                        return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);


                    }
                    else if (attendeeEmails.Count == 1)  // attendee found
                    {
                        attendeeTableStorage.Add(attendeeEmails[0]);
                    }
                    else //attendee not found in organization
                    {
                        await stepContext.Context.SendActivityAsync("Attendee not found, please type anything to start again");
                        return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
                    }

                }

                var sb = new System.Text.StringBuilder();
                foreach (string email in attendeeTableStorage)
                {
                    sb.Append(email + ","); //converts emails to comma separated string to store in table
                }
                string finalString = sb.ToString().Remove(sb.Length - 1);
                if (result != null)
                {
                    MeetingDetail meetingDetail = new MeetingDetail(stepContext.Context.Activity.Conversation.Id, userEmail);
                    meetingDetail.Attendees = finalString;

                    await InsertOrMergeEntityAsync(table, meetingDetail); //inserts attendees' emails in table

                    return await stepContext.PromptAsync(nameof(NumberPrompt<double>), new PromptOptions { Prompt = MessageFactory.Text("What will be duration of the meeting? (in hours)"), RetryPrompt = MessageFactory.Text("Invalid value, please enter a proper value") }, cancellationToken);
                }
            }
            await stepContext.Context.SendActivityAsync("Something went wrong. Please type anything to get started again.");

            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }
        private Task<bool> DurationPromptValidatorAsync(PromptValidatorContext<double> promptContext, CancellationToken cancellationToken)
        {
            return Task.FromResult(promptContext.Recognized.Succeeded && promptContext.Recognized.Value > 0 && promptContext.Recognized.Value < 8);

        }
 
        private async Task<DialogTurnResult> ShowMeetingTimeSuggestions(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            double duration= Convert.ToDouble(stepContext.Context.TurnState["data"]);
            var tokenResponse = (TokenResponse)stepContext.Result;
            if (tokenResponse?.Token != null)
            {
                // Pull in the data from the Microsoft Graph.
                var client = new SimpleGraphClient(tokenResponse.Token);
                MeetingDetail meetingDetail = await RetrieveMeetingDetailsAsync(table, userEmail, stepContext.Context.Activity.Conversation.Id); //retrives data from table

                timeSuggestions = await client.FindMeetingTimes(meetingDetail.Attendees, duration); //returns meeting times

                if (timeSuggestions.Count == 0)
                {
                    await stepContext.Context.SendActivityAsync("No appropriate meeting slot found. Please try again by typing 'hi' and change date this time");
                    return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);

                }
                var cardOptions = new List<Choice>();
                for (int i = 0; i < timeSuggestions.Count; i++)
                {
                    cardOptions.Add(new Choice() { Value = timeSuggestions[i].Start.DateTime + " - " + timeSuggestions[i].End.DateTime }); //creates list of meeting time choices
                }



                return await stepContext.PromptAsync(nameof(ChoicePrompt), new PromptOptions
                {
                    Prompt = MessageFactory.Text("These are the time suggestions. Click on the time slot for when you want the meeting to be set."),
                    RetryPrompt = MessageFactory.Text("Sorry, Please the valid choice"),
                    Choices = cardOptions,
                    Style = ListStyle.HeroCard, //displays choices as buttons
                }, cancellationToken);


            }
            await stepContext.Context.SendActivityAsync("Something went wrong. Please type anything to get started again.");
            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }

        private async Task<DialogTurnResult> AskForTitle(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var result = (FoundChoice)stepContext.Result;

            if (result != null)
            {
                MeetingDetail meetingDetail = new MeetingDetail(stepContext.Context.Activity.Conversation.Id, userEmail);
                meetingDetail.TimeSlotChoice = result.Index.ToString();

                await InsertOrMergeEntityAsync(table, meetingDetail); //inserts selected time slot in table

                return await stepContext.PromptAsync(nameof(TextPrompt), new PromptOptions { Prompt = MessageFactory.Text("Please enter title of the meeting") }, cancellationToken);

            }
            await stepContext.Context.SendActivityAsync("Something went wrong. Please type anything to get started again.");

            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }
        private async Task<DialogTurnResult> AskForDescription(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var result = (string)stepContext.Result;

            if (result != null)
            {
                MeetingDetail meetingDetail = new MeetingDetail(stepContext.Context.Activity.Conversation.Id, userEmail);
                meetingDetail.Title = result;

                await InsertOrMergeEntityAsync(table, meetingDetail); //inserts title in table

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
                MeetingDetail meetingDetail = await RetrieveMeetingDetailsAsync(table, userEmail, stepContext.Context.Activity.Conversation.Id); //retrieves current meeting details

                await client.SendMeetingInviteAsync(timeSuggestions[Int32.Parse(meetingDetail.TimeSlotChoice)], meetingDetail.Attendees, meetingDetail.Title, description); //creates event 

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
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(configuration["StorageConnectionString"]);
          

            // Create a table client for interacting with the table service
            CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());

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
                MeetingDetail meetingDetail = result.Result as MeetingDetail;

                // Get the request units consumed by the current operation. RequestCharge of a TableResult is only applied to Azure Cosmos DB


                return meetingDetail;
            }
            catch (Microsoft.Azure.Cosmos.Table.StorageException e)
            {

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

                throw;
            }
        }
        #endregion
    }
}
