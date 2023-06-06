using Microsoft.Bot.Schema;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using Microsoft.BotBuilderSamples.Models.dto;

namespace Microsoft.BotBuilderSamples.Models
{
    public class BatchConversationRequest
    {
        public List<ChannelAccount> Members { get; set; }
        public JToken Activity { get; set; }
        public string TenantId { get; set; }
        public string TeamId { get; set; }

        public BatchConversationRequestDto Convert()
        {
            return new BatchConversationRequestDto
            {
                Members = this.Members,
                Activity = this.Activity,
                TenantId = this.TenantId,
                TeamId = this.TeamId
            };
        }
    }
}
