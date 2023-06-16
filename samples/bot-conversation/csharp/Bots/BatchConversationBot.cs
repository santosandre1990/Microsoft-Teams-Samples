using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Schema;
using Microsoft.BotBuilderSamples.Models;
using Microsoft.BotBuilderSamples.Models.dto;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.BotBuilderSamples.Bots
{
    public class BatchConversationBot : TeamsActivityHandler
    {
        private readonly HttpClient httpClient;

        private struct CommandParameters
        {
            public ITurnContext<IMessageActivity> TurnContext;
            public IActivity Activity;
            public string Content;
            public IDictionary<string, string> Arguments;
            public int MentionTypes;
        };

        private delegate Task MessageCommandHandler(CommandParameters parameters);
        private readonly Dictionary<string, MessageCommandHandler> MessageCommandHandlers;

        public BatchConversationBot()
        {
            // Use canary endpoint
            httpClient = new HttpClient();
            httpClient.BaseAddress = new Uri("https://canary.botapi.skype.com/amer-df/");
            httpClient.Timeout = TimeSpan.FromSeconds(15);

            this.MessageCommandHandlers = new Dictionary<string, MessageCommandHandler>
            {
                { "all_tenant_users", SendMessageToAllTenantUsers },
                { "all_team_users", SendMessageToAllTeamUsers },
                { "list_of_channel", SendMessageToListOfChannels },
                { "list_of_enc_user_mri", SendMessageToListOfUsers },
                { "list_of_user_aad_obj_id", SendMessageToListOfUsers },
                { "cancel_operation", CancelOperation }
            };

        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            // Command sample
            //
            // audience -> "all_tenant_users","all_team_users","list_of_channel","list_of_enc_user_mri","list_of_user_aad_obj_id,"cancel_operation",
            // recipients argument valid for operations -> "list_of_channel","list_of_enc_user_mri","list_of_user_aad_obj_id"
            //
            // /batch_conversation -tenantId <test-tenant-id> -audience <test-audience> -content <test-content> -recipients <test-recipient-1>,<test-recipient-2>
            
            await ProcessMessageAsync(turnContext).ConfigureAwait(false);
        }

        private async Task SendMessageToListOfUsers(CommandParameters parameters)
        {

            if (!parameters.Arguments.ContainsKey("tenantId") || !parameters.Arguments.ContainsKey("recipients"))
            {
                await parameters.TurnContext.SendActivityAsync(MessageFactory.Text("Required parameters (tenantId, recipients) was not provided.")).ConfigureAwait(false);
                return;
            }

            var recipients = parameters.Arguments["recipients"];
            var entryIds = getEntryIds(recipients);

            BatchConversationRequest request = new BatchConversationRequest();
            request.Activity = JToken.FromObject(parameters.TurnContext.Activity);
            request.TenantId = parameters.TurnContext.Activity.Conversation.TenantId;
            request.Members = entryIds.Select(o => new ChannelAccount { Id = o }).ToList();

            // Create Async Batch Operation
            var createOperationResp = await postBatchMessagesAsync(parameters.TurnContext, request, BatchConversationEndpointType.listOfUsersEndpoint, CancellationToken.None).ConfigureAwait(false);

            // Wait for operation to complete
            var operationStateResp = await waitForOperationToComplete(parameters.TurnContext, createOperationResp.OperationId, CancellationToken.None);

            if (operationStateResp.StatusMap.Keys.Any(key => key != "201"))
            {
                // Check for failed entries - fetch first page - use continuation token to fetch more pages
                var failedEntriesPaginatedResp = await getFailedEntriesPaginatedAsync(parameters.TurnContext, createOperationResp.OperationId, CancellationToken.None);
            }
        }

        private async Task SendMessageToListOfChannels(CommandParameters parameters)
        {
            if (!parameters.Arguments.ContainsKey("tenantId") || !parameters.Arguments.ContainsKey("recipients"))
            {
                await parameters.TurnContext.SendActivityAsync(MessageFactory.Text("Required parameters (tenantId, recipients) was not provided.")).ConfigureAwait(false);
                return;
            }

            var recipients = parameters.Arguments["recipients"];
            var entryIds = getEntryIds(recipients);

            BatchConversationRequest request = new BatchConversationRequest();
            request.Activity = JToken.FromObject(parameters.TurnContext.Activity);
            request.TenantId = parameters.TurnContext.Activity.Conversation.TenantId;
            request.Members = entryIds.Select(o => new ChannelAccount { Id = o }).ToList();

            // Create Async Batch Operation
            var createOperationResp = await postBatchMessagesAsync(parameters.TurnContext, request, BatchConversationEndpointType.listOfChannelsEndpoint, CancellationToken.None).ConfigureAwait(false);

            // Wait for operation to complete
            var operationStateResp = await waitForOperationToComplete(parameters.TurnContext, createOperationResp.OperationId, CancellationToken.None);

            if (operationStateResp.StatusMap.Keys.Any(key => key != "201"))
            {
                // Check for failed entries - fetch first page - use continuation token to fetch more pages
                var failedEntriesPaginatedResp = await getFailedEntriesPaginatedAsync(parameters.TurnContext, createOperationResp.OperationId, CancellationToken.None);
            }
        }

        private async Task SendMessageToAllTenantUsers(CommandParameters parameters)
        {
            if (!parameters.Arguments.ContainsKey("tenantId"))
            {
                await parameters.TurnContext.SendActivityAsync(MessageFactory.Text("Required parameters (tenantId) was not provided.")).ConfigureAwait(false);
                return;
            }

            BatchConversationRequest request = new BatchConversationRequest();
            request.Activity = JToken.FromObject(parameters.TurnContext.Activity);
            request.TenantId = parameters.TurnContext.Activity.Conversation.TenantId;

            // Create Async Batch Operation
            var createOperationResp = await postBatchMessagesAsync(parameters.TurnContext, request, BatchConversationEndpointType.tenantUsersEndpoint, CancellationToken.None).ConfigureAwait(false);

            // Wait for operation to complete
            var operationStateResp = await waitForOperationToComplete(parameters.TurnContext, createOperationResp.OperationId, CancellationToken.None);

            if (operationStateResp.StatusMap.Keys.Any(key => key != "201"))
            {
                // Check for failed entries - fetch first page - use continuation token to fetch more pages
                var failedEntriesPaginatedResp = await getFailedEntriesPaginatedAsync(parameters.TurnContext, createOperationResp.OperationId, CancellationToken.None);
            }
        }

        private async Task SendMessageToAllTeamUsers(CommandParameters parameters)
        {

            if (!parameters.Arguments.ContainsKey("tenantId") || !parameters.Arguments.ContainsKey("teamId"))
            {
                await parameters.TurnContext.SendActivityAsync(MessageFactory.Text("Required parameters (tenantId,teamId) was not provided.")).ConfigureAwait(false);
                return;
            }

            BatchConversationRequest request = new BatchConversationRequest();
            request.Activity = JToken.FromObject(parameters.TurnContext.Activity);
            request.TenantId = parameters.TurnContext.Activity.Conversation.TenantId;
            request.TeamId = parameters.Arguments["teamId"];

            // Create Async Batch Operation
            var createOperationResp = await postBatchMessagesAsync(parameters.TurnContext, request, BatchConversationEndpointType.teamUserEndpoint, CancellationToken.None).ConfigureAwait(false);

            // Wait for operation to complete
            var operationStateResp = await waitForOperationToComplete(parameters.TurnContext, createOperationResp.OperationId, CancellationToken.None);

            if (operationStateResp.StatusMap.Keys.Any(key => key != "201"))
            {
                // Check for failed entries - fetch first page - use continuation token to fetch more pages
                var failedEntriesPaginatedResp = await getFailedEntriesPaginatedAsync(parameters.TurnContext, createOperationResp.OperationId, CancellationToken.None);
            }
        }

        private async Task CancelOperation(CommandParameters parameters)
        {

            if (!parameters.Arguments.ContainsKey("operationId"))
            {
                await parameters.TurnContext.SendActivityAsync(MessageFactory.Text("Required parameters (operationId) was not provided.")).ConfigureAwait(false);
                return;
            }

            // Cancel Operation
            var operationId = parameters.Arguments["operationId"];
            
            await cancelOperationAsync(parameters.TurnContext, operationId, CancellationToken.None).ConfigureAwait(false);
        }

        private async Task<GetBatchConversationStateResponse> waitForOperationToComplete(ITurnContext<IMessageActivity> turnContext, string operationId, CancellationToken cancellationToken)
        {
            GetBatchConversationStateResponse getOpStateResp;
            do
            {
                // Get Operation State
                getOpStateResp = await getBatchOperationStateAsync(turnContext, operationId, cancellationToken);

                // Operation is ongoing while state in "Ongoing" or "Provisioning"
                if (getOpStateResp.State.ToLower() == "ongoing" || getOpStateResp.State.ToLower() == "provisioning")
                {
                    // Retries should respect the Retry-After property value, or else the bot will be throttled
                    var t = getOpStateResp.RetryAfter.Value.Subtract(DateTime.UtcNow);

                    await Task.Delay((t.Seconds + t.Minutes * 60) * 1000);
                    continue;

                }
            } while (getOpStateResp.State.ToLower() != "completed" && getOpStateResp.State.ToLower() != "failed" || cancellationToken.IsCancellationRequested);

            return getOpStateResp;
        }

        #region HTTP client helpers - Valid until SDK includes new APIs

        private async Task<CreateBatchConversationResponse> postBatchMessagesAsync(
            ITurnContext<IMessageActivity> turnContext,
            BatchConversationRequest requestBody,
            BatchConversationEndpointType endpoint,
            CancellationToken cancellationToken
            )
        {

            var Client = turnContext.TurnState.Get<IConnectorClient>();
            var creds = Client.Credentials as AppCredentials;
            var bearerToken = await creds.GetTokenAsync().ConfigureAwait(false);

            using (var request = new HttpRequestMessage())
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                request.RequestUri = new Uri(mapBatchConversationApiEndpoints(endpoint), UriKind.Relative);
                request.Method = HttpMethod.Post;
                request.Content = new StringContent(requestBody.Convert().Serialize(), Encoding.UTF8, "application/json");

                using (HttpResponseMessage response = await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    string content = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception();
                    }

                    return CreateBatchConversationResponseDto.Deserialize(content);
                }
            }
        }

        private async Task<GetBatchConversationStateResponse> getBatchOperationStateAsync(
            ITurnContext<IMessageActivity> turnContext,
            string operationId,
            CancellationToken cancellationToken
            )
        {

            var Client = turnContext.TurnState.Get<IConnectorClient>();
            var creds = Client.Credentials as AppCredentials;
            var bearerToken = await creds.GetTokenAsync().ConfigureAwait(false);

            using (var request = new HttpRequestMessage())
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                request.RequestUri = new Uri(string.Concat(mapBatchConversationApiEndpoints(BatchConversationEndpointType.operationState), operationId), UriKind.Relative);
                request.Method = HttpMethod.Get;

                using (HttpResponseMessage response = await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    string content = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception();
                    }

                    return GetBatchConversationStateResponseDto.Deserialize(content);
                }
            }
        }

        private async Task<GetFailedEntriesResponse> getFailedEntriesPaginatedAsync(
            ITurnContext<IMessageActivity> turnContext,
            string operationId,
            CancellationToken cancellationToken,
            string continuationToken = null
            )
        {

            var Client = turnContext.TurnState.Get<IConnectorClient>();
            var creds = Client.Credentials as AppCredentials;
            var bearerToken = await creds.GetTokenAsync().ConfigureAwait(false);

            using (var request = new HttpRequestMessage())
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);

                if (continuationToken != null)
                    request.RequestUri = new Uri(string.Concat(mapBatchConversationApiEndpoints(BatchConversationEndpointType.failedEntriesPaginated), operationId, $"?continuationToken={continuationToken}"), UriKind.Relative);
                else
                    request.RequestUri = new Uri(string.Concat(mapBatchConversationApiEndpoints(BatchConversationEndpointType.failedEntriesPaginated), operationId), UriKind.Relative);

                request.Method = HttpMethod.Get;

                using (HttpResponseMessage response = await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    string content = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception();
                    }

                    return GetFailedEntriesResponseDto.Deserialize(content);
                }
            }
        }

        private async Task<bool> cancelOperationAsync(
            ITurnContext<IMessageActivity> turnContext,
            string operationId,
            CancellationToken cancellationToken
            )
        {

            var Client = turnContext.TurnState.Get<IConnectorClient>();
            var creds = Client.Credentials as AppCredentials;
            var bearerToken = await creds.GetTokenAsync().ConfigureAwait(false);

            using (var request = new HttpRequestMessage())
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                request.RequestUri = new Uri(string.Concat(mapBatchConversationApiEndpoints(BatchConversationEndpointType.cancelOperation), operationId), UriKind.Relative);
                request.Method = HttpMethod.Delete;

                using (HttpResponseMessage response = await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    string content = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                        return false;

                    return true;
                }
            }
        }

        private string mapBatchConversationApiEndpoints(BatchConversationEndpointType endpoint)
        {
            switch (endpoint)
            {
                case BatchConversationEndpointType.listOfChannelsEndpoint:
                    return "v3/batch/conversation/channels";
                case BatchConversationEndpointType.tenantUsersEndpoint:
                    return "v3/batch/conversation/tenant";
                case BatchConversationEndpointType.teamUserEndpoint:
                    return "v3/batch/conversation/team";
                case BatchConversationEndpointType.listOfUsersEndpoint:
                    return "v3/batch/conversation/users";
                case BatchConversationEndpointType.operationState:
                    return "v3/batch/conversation/";
                case BatchConversationEndpointType.failedEntriesPaginated:
                    return "v3/batch/conversation/failedentries/";
                case BatchConversationEndpointType.cancelOperation:
                    return "v3/batch/conversation/";
                default:
                    throw new Exception($"Provided endpoint is not valid");
            }
        }
        #endregion

        #region Helpers

        private List<string> getEntryIds(string recipients)
        {
            var recipientList = recipients.Split(",").ToList();
            List<string> listOfEntriesIds = new List<string>();
            foreach (string recipient in recipientList)
            {
                listOfEntriesIds.Add(recipient);
            }
            return listOfEntriesIds;
        }

        private async Task ProcessMessageAsync(ITurnContext<IMessageActivity> turnContext)
        {
            var parameters = new CommandParameters()
            {
                TurnContext = turnContext,
                Activity = turnContext.Activity,
            };

            string content = turnContext.Activity.Text;

            List<string> splitted = content.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToList();
           
            // Parse arguments
            IDictionary<string, string> arguments = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            int i = 1;

            while (i < splitted.Count())
            {
                if (!splitted[i].StartsWith("-") || splitted[i].Equals("-") || i + 1 >= splitted.Count() || splitted[i + 1].StartsWith("-"))
                {
                    await SendMessageAsync(turnContext, "Invalid command format.");

                    return;
                }

                arguments.Add(splitted[i].Substring(1), splitted[i + 1]);
                i += 2;
            }

            if (!arguments.ContainsKey("audience"))
            {
                await SendMessageAsync(turnContext, "Invalid command format. Missing audience argument.");
                return;
            }

            var methodPair = MessageCommandHandlers.Where(t => string.Equals(arguments["audience"], t.Key, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

            // Not found
            if (string.IsNullOrWhiteSpace(methodPair.Key))
            {
                await SendMessageAsync(turnContext, "Invalid command format.");
                return;
            }

            try
            {
                parameters.Content = content.Substring(methodPair.Key.Length).Trim();
                parameters.Arguments = arguments;

                await methodPair.Value(parameters).ConfigureAwait(false);
            }
            catch (Exception e)
            {
                var replyToConversation = MessageFactory.Text(e.ToString());
                await parameters.TurnContext.SendActivityAsync(replyToConversation).ConfigureAwait(false);
            }
        }

        public static async Task SendMessageAsync<T>(ITurnContext<T> turnContext, string content) where T : IActivity
        {
            var replyToConversation = MessageFactory.Text(content);
            replyToConversation.TextFormat = (turnContext.Activity as IMessageActivity).TextFormat;
            await turnContext.SendActivityAsync(replyToConversation).ConfigureAwait(false);
        }

        #endregion
    }
}
