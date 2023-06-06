using Microsoft.BotBuilderSamples.Models;
using Newtonsoft.Json;
using System.Collections.Generic;
using System;

namespace Microsoft.BotBuilderSamples.Models.dto
{
    public class GetBatchConversationStateResponseDto
    {
        [JsonProperty("State")]
        public string State { get; set; }

        [JsonProperty("RetryAfter")]
        public DateTime? RetryAfter { get; set; } = null;

        [JsonProperty("StatusMap")]
        public Dictionary<string, int> StatusMap { get; set; }

        [JsonProperty("TotalUserCount")]
        public int TotalUserCount { get; set; }

        public static GetBatchConversationStateResponse Deserialize(string json)
        {
            return JsonConvert.DeserializeObject<GetBatchConversationStateResponse>(json);
        }

    }
}
