using System.Collections.Generic;
using System;

namespace Microsoft.BotBuilderSamples.Models
{
    public class GetBatchConversationStateResponse
    {
        public string State { get; set; }
        public DateTime? RetryAfter { get; set; } = null;
        public Dictionary<string, int> StatusMap { get; set; }
        public int TotalUserCount { get; set; }
    }
}
