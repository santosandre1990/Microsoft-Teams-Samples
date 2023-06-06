using Newtonsoft.Json;
using System;
using System.Collections.Generic;

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
