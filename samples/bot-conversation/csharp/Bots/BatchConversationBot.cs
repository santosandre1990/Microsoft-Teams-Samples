using Microsoft.Bot.Builder;
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
using static Microsoft.ApplicationInsights.MetricDimensionNames.TelemetryContext;

namespace Microsoft.BotBuilderSamples.Bots
{
    public class BatchConversationBot
    {
        private readonly HttpClient httpClient;

        public BatchConversationBot()
        {
            httpClient = new HttpClient();
            // Use canary endpoint
            httpClient.BaseAddress = new Uri("https://canary.botapi.skype.com/amer-df/");
            httpClient.Timeout = TimeSpan.FromSeconds(15);
        }

        public async Task SendMessageToListOfUsers(ITurnContext<IMessageActivity> turnContext, string tenantId, List<string> users, CancellationToken cancellationToken)
        {
            BatchConversationRequest request = new BatchConversationRequest();
            request.Activity = JToken.FromObject(turnContext.Activity);
            request.TenantId = tenantId;
            request.Members = users.Select(o => new ChannelAccount { Id = o }).ToList();

            var createOperationResp = await sendBatchMessagesAsync(turnContext, request, BatchConversationEndpointType.listOfUsersEndpoint, cancellationToken).ConfigureAwait(false);

            // Wait for operation to complete
            await waitForOperationState(turnContext, createOperationResp.OperationId, cancellationToken);
        }

        public async Task SendMessageToListOfChannels(ITurnContext<IMessageActivity> turnContext, string tenantId, List<string> channels, CancellationToken cancellationToken)
        {
            BatchConversationRequest request = new BatchConversationRequest();
            request.Activity = JToken.FromObject(turnContext.Activity);
            request.TenantId = tenantId;
            request.Members = channels.Select(o => new ChannelAccount { Id = o }).ToList();

            var createOperationResp = await sendBatchMessagesAsync(turnContext, request, BatchConversationEndpointType.listOfChannelsEndpoint, cancellationToken).ConfigureAwait(false);

            // Wait for operation to complete
            await waitForOperationState(turnContext, createOperationResp.OperationId, cancellationToken);
        }

        public async Task SendMessageToAllTenantUsers(ITurnContext<IMessageActivity> turnContext, string tenantId, CancellationToken cancellationToken)
        {
            BatchConversationRequest request = new BatchConversationRequest();
            request.Activity = JToken.FromObject(turnContext.Activity);
            request.TenantId = tenantId;

            var createOperationResp = await sendBatchMessagesAsync(turnContext, request, BatchConversationEndpointType.listOfUsersEndpoint, cancellationToken).ConfigureAwait(false);

            // Wait for operation to complete
            await waitForOperationState(turnContext, createOperationResp.OperationId, cancellationToken);
        }

        public async Task SendMessageToTeamsUsers(ITurnContext<IMessageActivity> turnContext, string tenantId, string teamId, CancellationToken cancellationToken)
        {
            BatchConversationRequest request = new BatchConversationRequest();
            request.Activity = JToken.FromObject(turnContext.Activity);
            request.TenantId = tenantId;
            request.TeamId = teamId;

            var createOperationResp = await sendBatchMessagesAsync(turnContext, request, BatchConversationEndpointType.listOfUsersEndpoint, cancellationToken).ConfigureAwait(false);

            // Wait for operation to complete
            await waitForOperationState(turnContext, createOperationResp.OperationId, cancellationToken);
        }

        private async Task<GetBatchConversationStateResponse> waitForOperationState(ITurnContext<IMessageActivity> turnContext, string operationId, CancellationToken cancellationToken)
        {
            GetBatchConversationStateResponse getOpStateResp;
            do
            {
                // Fetch Retry after
                getOpStateResp = await getBatchOperationStateAsync(turnContext, operationId, cancellationToken);

                if (getOpStateResp.State.ToLower() == "ongoing" || getOpStateResp.State.ToLower() == "provisioning")
                {
                    var t = getOpStateResp.RetryAfter.Value.Subtract(DateTime.UtcNow);

                        await Task.Delay((t.Seconds + t.Minutes * 60) * 1000);
                        continue;

                }
            } while (getOpStateResp.State.ToLower() != "completed" && getOpStateResp.State.ToLower() != "failed" || cancellationToken.IsCancellationRequested);

            return getOpStateResp;
        }

        #region HTTP client helpers - Valid until SDK includes new APIs

        public async Task<CreateBatchConversationResponse> sendBatchMessagesAsync(
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
                request.RequestUri = new Uri(string.Concat("v3/batch/conversation/", operationId), UriKind.Relative);
                request.Method = HttpMethod.Get;

                using (HttpResponseMessage response = await httpClient.SendAsync(request,cancellationToken).ConfigureAwait(false))
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
                default:
                    throw new Exception($"Provided endpoint is not valid");
            }
        }
        #endregion
    }
}
