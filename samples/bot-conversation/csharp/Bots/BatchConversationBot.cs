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

        public BatchConversationBot()
        {
            httpClient = new HttpClient();
            // Use canary endpoint
            httpClient.BaseAddress = new Uri("https://canary.botapi.skype.com/amer-df/");
            httpClient.Timeout = TimeSpan.FromSeconds(15);
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            turnContext.Activity.RemoveRecipientMention();
            var text = turnContext.Activity.Text.Trim().ToLower();

            //if (text.Contains("tenant"))

            //else if (text.Contains("team"))

            //else if (text.Contains("channel"))

            //else if (text.Contains("users"))

            //else
            //    throw new NotImplementedException();
        }

        public async Task SendMessageToListOfUsers(ITurnContext<IMessageActivity> turnContext, string tenantId, List<string> users, CancellationToken cancellationToken)
        {
            BatchConversationRequest request = new BatchConversationRequest();
            request.Activity = JToken.FromObject(turnContext.Activity);
            request.TenantId = tenantId;
            request.Members = users.Select(o => new ChannelAccount { Id = o }).ToList();

            // Create Async Batch Operation
            var createOperationResp = await postBatchMessagesAsync(turnContext, request, BatchConversationEndpointType.listOfUsersEndpoint, cancellationToken).ConfigureAwait(false);

            // Wait for operation to complete
            var operationStateResp = await waitForOperationToComplete(turnContext, createOperationResp.OperationId, cancellationToken);

            if(operationStateResp.StatusMap.Keys.Any(key => key != "201"))
            {
                // Check for failed entries - fetch first page - use continuation token to fetch more pages
                var failedEntriesPaginatedResp = await getFailedEntriesPaginatedAsync(turnContext, createOperationResp.OperationId, cancellationToken);
            }
        }

        public async Task SendMessageToListOfChannels(ITurnContext<IMessageActivity> turnContext, string tenantId, List<string> channels, CancellationToken cancellationToken)
        {
            BatchConversationRequest request = new BatchConversationRequest();
            request.Activity = JToken.FromObject(turnContext.Activity);
            request.TenantId = tenantId;
            request.Members = channels.Select(o => new ChannelAccount { Id = o }).ToList();

            // Create Async Batch Operation
            var createOperationResp = await postBatchMessagesAsync(turnContext, request, BatchConversationEndpointType.listOfChannelsEndpoint, cancellationToken).ConfigureAwait(false);

            // Wait for operation to complete
            var operationStateResp = await waitForOperationToComplete(turnContext, createOperationResp.OperationId, cancellationToken);

            if (operationStateResp.StatusMap.Keys.Any(key => key != "201"))
            {
                // Check for failed entries - fetch first page - use continuation token to fetch more pages
                var failedEntriesPaginatedResp = await getFailedEntriesPaginatedAsync(turnContext, createOperationResp.OperationId, cancellationToken);
            }
        }

        public async Task SendMessageToAllTenantUsers(ITurnContext<IMessageActivity> turnContext, string tenantId, CancellationToken cancellationToken)
        {
            BatchConversationRequest request = new BatchConversationRequest();
            request.Activity = JToken.FromObject(turnContext.Activity);
            request.TenantId = tenantId;

            // Create Async Batch Operation
            var createOperationResp = await postBatchMessagesAsync(turnContext, request, BatchConversationEndpointType.tenantUsersEndpoint, cancellationToken).ConfigureAwait(false);

            // Wait for operation to complete
            var operationStateResp = await waitForOperationToComplete(turnContext, createOperationResp.OperationId, cancellationToken);

            if (operationStateResp.StatusMap.Keys.Any(key => key != "201"))
            {
                // Check for failed entries - fetch first page - use continuation token to fetch more pages
                var failedEntriesPaginatedResp = await getFailedEntriesPaginatedAsync(turnContext, createOperationResp.OperationId, cancellationToken);
            }
        }

        public async Task SendMessageToTeamsUsers(ITurnContext<IMessageActivity> turnContext, string tenantId, string teamId, CancellationToken cancellationToken)
        {
            BatchConversationRequest request = new BatchConversationRequest();
            request.Activity = JToken.FromObject(turnContext.Activity);
            request.TenantId = tenantId;
            request.TeamId = teamId;

            // Create Async Batch Operation
            var createOperationResp = await postBatchMessagesAsync(turnContext, request, BatchConversationEndpointType.teamUserEndpoint, cancellationToken).ConfigureAwait(false);

            // Wait for operation to complete
            var operationStateResp = await waitForOperationToComplete(turnContext, createOperationResp.OperationId, cancellationToken);

            if (operationStateResp.StatusMap.Keys.Any(key => key != "201"))
            {
                // Check for failed entries - fetch first page - use continuation token to fetch more pages
                var failedEntriesPaginatedResp = await getFailedEntriesPaginatedAsync(turnContext, createOperationResp.OperationId, cancellationToken);
            }
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

        public async Task<CreateBatchConversationResponse> postBatchMessagesAsync(
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

        public async Task<GetBatchConversationStateResponse> getBatchOperationStateAsync(
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

        public async Task<GetFailedEntriesResponse> getFailedEntriesPaginatedAsync(
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
                
                if(continuationToken != null )
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

        public async Task<bool> cancelOperationAsync(
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
    }
}
